"""
Kaspi HTTP client — no Selenium, no browser.
Uses public Kaspi JSON API directly.
"""

import asyncio
import re
from typing import Awaitable, Callable, Optional

import httpx
from bs4 import BeautifulSoup

BASE_URL = "https://kaspi.kz"

HEADERS = {
    "Accept": "application/json, text/*",
    "Accept-Language": "ru,en;q=0.9",
    "Content-Type": "application/json; charset=UTF-8",
    "Origin": "https://kaspi.kz",
    "Referer": "https://kaspi.kz/",
    "Sec-Fetch-Dest": "empty",
    "Sec-Fetch-Mode": "cors",
    "Sec-Fetch-Site": "same-origin",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/144.0.0.0 Safari/537.36"
    ),
    "X-Description-Enabled": "true",
}

CITY_SLUGS = {
    "710000000": "astana",
    "551010000": "pavlodar",
}


def _sku_from_href(href: str) -> Optional[str]:
    """Extract masterSku from product URL like /shop/p/some-name-134842334/"""
    m = re.search(r"-(\d{6,12})/?(?:\?|$)", href)
    return m.group(1) if m else None


def _parse_product_cards(html: str) -> list[dict]:
    """Extract {sku, name, url} from category page HTML."""
    soup = BeautifulSoup(html, "lxml")
    items = []
    seen: set[str] = set()

    for a in soup.select("a.item-card__name-link"):
        href = a.get("href", "")
        sku = _sku_from_href(href)
        if not sku or sku in seen:
            continue
        seen.add(sku)
        url = ("https://kaspi.kz" + href.split("?")[0]) if href.startswith("/") else href.split("?")[0]
        items.append({"sku": sku, "name": a.get_text(strip=True), "url": url})

    return items


def _total_pages(html: str) -> int:
    """Extract last page number from category HTML pagination."""
    soup = BeautifulSoup(html, "lxml")

    # Try numbered pagination links
    nums = [
        int(a.get_text())
        for a in soup.select("li.pagination__el a")
        if a.get_text(strip=True).isdigit()
    ]
    if nums:
        return max(nums)

    # If there's a "next" button, at least one more page exists
    if soup.select_one("li.pagination__el--next:not(.disabled)"):
        return 999  # will break naturally when no items returned

    return 1


class KaspiClient:
    def __init__(self, concurrency: int = 15):
        self._sem = asyncio.Semaphore(concurrency)
        self._client: Optional[httpx.AsyncClient] = None

    async def __aenter__(self):
        self._client = httpx.AsyncClient(
            base_url=BASE_URL,
            headers=HEADERS,
            timeout=30.0,
            follow_redirects=True,
        )
        return self

    async def __aexit__(self, *_):
        if self._client:
            await self._client.aclose()

    # ── Low-level ──────────────────────────────────────────────────────────────

    async def _get(self, path: str, params: dict = None) -> str:
        async with self._sem:
            r = await self._client.get(path, params=params)
            r.raise_for_status()
            return r.text

    async def _post(self, path: str, body: dict, city_id: str) -> dict | list:
        async with self._sem:
            r = await self._client.post(
                path,
                json=body,
                headers={"X-KS-City": city_id},
            )
            r.raise_for_status()
            return r.json()

    # ── Category scraping ──────────────────────────────────────────────────────

    async def get_category_skus(
        self,
        category: str,
        city_id: str,
        merchant_id: Optional[str] = None,
        extra_filters: Optional[dict] = None,
        on_page: Optional[Callable[[int, int], Awaitable[None]]] = None,
    ) -> list[dict]:
        """
        Scrape all pages of a category and return list of {sku, name, url}.

        category:      e.g. "rims", "wheels", "auto%20parts"
        city_id:       "710000000" (Astana) or "551010000" (Pavlodar)
        merchant_id:   filter by seller, e.g. "Kama"
        extra_filters: additional q-filters, e.g. {"Rims*Type": "литые"}
        """
        city_slug = CITY_SLUGS.get(city_id, "pavlodar")
        path = f"/shop/{city_slug}/c/{category}/"

        q_parts = [f":availableInZones:{city_id}", f":category:{category}"]
        if merchant_id:
            q_parts.append(f":allMerchants:{merchant_id}")
        if extra_filters:
            for k, v in extra_filters.items():
                q_parts.append(f":{k}:{v}")
        q = "".join(q_parts)

        all_items: list[dict] = []
        page = 0

        while True:
            params = {"q": q, "sort": "relevance"}
            if page > 0:
                params["page"] = page

            html = await self._get(path, params)
            items = _parse_product_cards(html)

            if not items:
                break

            all_items.extend(items)
            total_pages = _total_pages(html)

            if on_page:
                await on_page(page + 1, total_pages)

            page += 1
            if page >= total_pages:
                break

            await asyncio.sleep(0.4)  # polite rate-limiting

        # Final dedup (pages can overlap)
        seen: set[str] = set()
        result = []
        for item in all_items:
            if item["sku"] not in seen:
                seen.add(item["sku"])
                result.append(item)
        return result

    # ── Offers / pickup ────────────────────────────────────────────────────────

    async def check_pickup(
        self,
        sku: str,
        city_id: str,
        merchant_id: Optional[str] = None,
    ) -> Optional[dict]:
        """
        Returns the offer dict if PICKUP is available in the given city, else None.
        Optionally filter by a specific merchant.
        """
        body = {
            "cityId": city_id,
            "id": sku,
            "merchantUID": [merchant_id] if merchant_id else [],
            "limit": 20,
            "page": 0,
            "product": {
                "brand": "",
                "categoryCodes": [],
                "baseProductCodes": [],
                "groups": None,
                "productSeries": [],
            },
            "sortOption": "PRICE",
            "highRating": None,
            "searchText": None,
            "isExcellentMerchant": False,
            "zoneId": [city_id],
            "installationId": "-1",
        }
        data = await self._post(f"/yml/offer-view/offers/{sku}", body, city_id)
        for offer in data.get("offers", []):
            if "PICKUP" in offer.get("deliveryOptions", {}):
                return offer
        return None

    async def bulk_check_pickup(
        self,
        items: list[dict],
        city_id: str,
        merchant_id: Optional[str] = None,
        on_progress: Optional[Callable[[int], Awaitable[None]]] = None,
    ) -> list[dict]:
        """
        Check pickup for a list of items concurrently.
        items: list of {sku, name, url}
        Returns list of {Код товара, Наименование, URL, Цена} for items with pickup.
        """
        results: list[dict] = []
        checked = 0
        lock = asyncio.Lock()

        async def _check(item: dict):
            nonlocal checked
            try:
                offer = await self.check_pickup(item["sku"], city_id, merchant_id)
                if offer:
                    async with lock:
                        results.append({
                            "Код товара": item["sku"],
                            "Наименование": offer.get("title") or item.get("name", ""),
                            "URL": item.get("url", f"https://kaspi.kz/shop/p/{item['sku']}/"),
                            "Цена": offer.get("price", ""),
                        })
            except Exception:
                pass
            finally:
                async with lock:
                    checked += 1
                if on_progress:
                    await on_progress(checked)

        await asyncio.gather(*[_check(item) for item in items])
        return results

    async def get_merchant_products_with_city(
        self,
        category: str,
        merchant_id: str,
        on_progress: Optional[Callable[[int, int], Awaitable[None]]] = None,
    ) -> list[dict]:
        """
        Get all products of a merchant and classify them by city (Pavlodar / Astana).
        Logic: if the product has PICKUP in Pavlodar → it's in Pavlodar, else Astana.
        Returns list of {Код товара, Наименование, URL, Город, Цена}.
        """
        items = await self.get_category_skus(
            category, "551010000", merchant_id=merchant_id
        )

        results: list[dict] = []
        lock = asyncio.Lock()
        processed = 0

        async def _check(item: dict, idx: int):
            nonlocal processed
            try:
                pvl_offer = await self.check_pickup(item["sku"], "551010000", merchant_id)
                city = "Павлодар" if pvl_offer else "Астана"
                price = pvl_offer.get("price", "") if pvl_offer else ""
                async with lock:
                    results.append({
                        "Код товара": item["sku"],
                        "Наименование": item.get("name", ""),
                        "URL": item.get("url", ""),
                        "Город": city,
                        "Цена": price,
                    })
            except Exception:
                pass
            finally:
                async with lock:
                    processed += 1
                if on_progress:
                    await on_progress(processed, len(items))

        await asyncio.gather(*[_check(item, i) for i, item in enumerate(items)])
        return results
