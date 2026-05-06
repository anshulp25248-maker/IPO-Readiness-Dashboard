from __future__ import annotations

import json
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import urlencode

import requests
from requests.adapters import HTTPAdapter
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver import ChromeOptions
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from urllib3.util.retry import Retry

from config import Settings, get_logger


@dataclass(slots=True)
class PostingResult:
    """Unified result object for both posting strategies."""

    success: bool
    method: str
    message: str
    post_id: str = ""


class LinkedInAPIPoster:
    """Posts directly through the official LinkedIn UGC API."""

    AUTH_URL = "https://www.linkedin.com/oauth/v2/authorization"
    TOKEN_URL = "https://www.linkedin.com/oauth/v2/accessToken"
    USERINFO_URL = "https://api.linkedin.com/v2/userinfo"
    PROFILE_URL = "https://api.linkedin.com/v2/me"
    UGC_POSTS_URL = "https://api.linkedin.com/v2/ugcPosts"

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("linkedin_api")
        self.session = self._build_session()

    def build_authorization_url(self, state: str = "linkedin-macro-agent") -> str:
        """Build the LinkedIn OAuth URL used for first-time API setup."""
        if not self.settings.linkedin_client_id or not self.settings.linkedin_redirect_uri:
            raise ValueError("LINKEDIN_CLIENT_ID and LINKEDIN_REDIRECT_URI are required to build the OAuth URL.")

        params = {
            "response_type": "code",
            "client_id": self.settings.linkedin_client_id,
            "redirect_uri": self.settings.linkedin_redirect_uri,
            "scope": "openid profile w_member_social",
            "state": state,
        }
        return f"{self.AUTH_URL}?{urlencode(params)}"

    def exchange_code_for_access_token(self, authorization_code: str) -> dict[str, Any]:
        """Exchange the temporary LinkedIn authorization code for a reusable access token."""
        required = (
            self.settings.linkedin_client_id,
            self.settings.linkedin_client_secret,
            self.settings.linkedin_redirect_uri,
        )
        if not all(required):
            raise ValueError(
                "LINKEDIN_CLIENT_ID, LINKEDIN_CLIENT_SECRET, and LINKEDIN_REDIRECT_URI are required to exchange a code."
            )

        payload = {
            "grant_type": "authorization_code",
            "code": authorization_code,
            "client_id": self.settings.linkedin_client_id,
            "client_secret": self.settings.linkedin_client_secret,
            "redirect_uri": self.settings.linkedin_redirect_uri,
        }
        response = self.session.post(
            self.TOKEN_URL,
            data=payload,
            timeout=self.settings.request_timeout_seconds,
        )
        response.raise_for_status()
        return response.json()

    def fetch_member_profile(self, access_token: str | None = None) -> dict[str, Any]:
        """Fetch the authenticated member profile so you can discover the person id."""
        token = access_token or self.settings.linkedin_access_token
        if not token:
            raise ValueError("LINKEDIN_ACCESS_TOKEN is required to fetch the LinkedIn profile.")

        response = self.session.get(
            self.USERINFO_URL,
            headers={"Authorization": f"Bearer {token}"},
            timeout=self.settings.request_timeout_seconds,
        )
        if response.ok:
            payload = response.json()
            if payload.get("sub"):
                payload["person_urn"] = f"urn:li:person:{payload['sub']}"
                return payload

        response = self.session.get(
            self.PROFILE_URL,
            headers={
                "Authorization": f"Bearer {token}",
                "X-Restli-Protocol-Version": "2.0.0",
            },
            timeout=self.settings.request_timeout_seconds,
        )
        response.raise_for_status()
        payload = response.json()
        if payload.get("id"):
            payload["person_urn"] = f"urn:li:person:{payload['id']}"
        return payload

    def post_text(self, post_text: str) -> PostingResult:
        """Create a LinkedIn text post through the API."""
        if not self.settings.linkedin_access_token:
            return PostingResult(False, "api", "LINKEDIN_ACCESS_TOKEN is missing.")

        try:
            author_urn = self._resolve_author_urn()
        except Exception as exc:
            return PostingResult(False, "api", f"Could not determine LinkedIn author URN: {exc}")

        payload = {
            "author": author_urn,
            "lifecycleState": "PUBLISHED",
            "specificContent": {
                "com.linkedin.ugc.ShareContent": {
                    "shareCommentary": {"text": post_text},
                    "shareMediaCategory": "NONE",
                }
            },
            "visibility": {
                "com.linkedin.ugc.MemberNetworkVisibility": "PUBLIC",
            },
        }

        response = self.session.post(
            self.UGC_POSTS_URL,
            headers={
                "Authorization": f"Bearer {self.settings.linkedin_access_token}",
                "X-Restli-Protocol-Version": "2.0.0",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=self.settings.request_timeout_seconds,
        )

        if not response.ok:
            message = self._build_error_message(response)
            return PostingResult(False, "api", message)

        post_id = response.headers.get("x-restli-id", "")
        return PostingResult(True, "api", "Post published through the LinkedIn API.", post_id=post_id)

    def _resolve_author_urn(self) -> str:
        """Resolve the member URN from env vars or a profile lookup."""
        if self.settings.linkedin_person_urn:
            return self.settings.linkedin_person_urn
        if self.settings.linkedin_person_id:
            return f"urn:li:person:{self.settings.linkedin_person_id}"

        profile = self.fetch_member_profile()
        if profile.get("person_urn"):
            return profile["person_urn"]
        if profile.get("sub"):
            return f"urn:li:person:{profile['sub']}"
        if profile.get("id"):
            return f"urn:li:person:{profile['id']}"

        raise ValueError(
            "No person id was found. Set LINKEDIN_PERSON_ID or LINKEDIN_PERSON_URN in your .env file."
        )

    def _build_session(self) -> requests.Session:
        """Retry transient API failures like rate limits or upstream errors."""
        session = requests.Session()
        retry = Retry(
            total=3,
            connect=3,
            read=3,
            backoff_factor=1.0,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods={"GET", "POST"},
        )
        adapter = HTTPAdapter(max_retries=retry)
        session.mount("https://", adapter)
        session.mount("http://", adapter)
        return session

    @staticmethod
    def _build_error_message(response: requests.Response) -> str:
        try:
            payload = response.json()
            return f"LinkedIn API error {response.status_code}: {json.dumps(payload, ensure_ascii=False)}"
        except ValueError:
            return f"LinkedIn API error {response.status_code}: {response.text}"


class LinkedInSeleniumPoster:
    """Fallback browser automation for local posting when the API path is unavailable."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("linkedin_selenium")

    def post_text(self, post_text: str) -> PostingResult:
        """Log in to LinkedIn in a browser and submit a new text post."""
        if not self.settings.linkedin_email or not self.settings.linkedin_password:
            return PostingResult(False, "selenium", "LINKEDIN_EMAIL and LINKEDIN_PASSWORD are required.")

        driver = None
        try:
            driver = self._build_driver()
            wait = WebDriverWait(driver, 30)

            self._login(driver, wait)
            self._create_post(driver, wait, post_text)
            return PostingResult(True, "selenium", "Post published through Selenium automation.")
        except Exception as exc:
            return PostingResult(False, "selenium", f"Selenium posting failed: {exc}")
        finally:
            if driver:
                driver.quit()

    def _build_driver(self) -> webdriver.Chrome:
        """Start Chrome with sensible local defaults."""
        options = ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")

        if self.settings.linkedin_headless:
            options.add_argument("--headless=new")

        if self.settings.chrome_binary_path:
            options.binary_location = self.settings.chrome_binary_path

        return webdriver.Chrome(options=options)

    def _login(self, driver: webdriver.Chrome, wait: WebDriverWait) -> None:
        """Perform the LinkedIn login flow."""
        driver.get("https://www.linkedin.com/login")
        wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(self.settings.linkedin_email)
        driver.find_element(By.ID, "password").send_keys(self.settings.linkedin_password)
        driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()

        wait.until(lambda current_driver: "feed" in current_driver.current_url or "checkpoint" in current_driver.current_url)
        if "checkpoint" in driver.current_url:
            raise RuntimeError("LinkedIn asked for a checkpoint, CAPTCHA, or 2FA challenge. Complete it manually first.")

    def _create_post(self, driver: webdriver.Chrome, wait: WebDriverWait, post_text: str) -> None:
        """Open the compose modal, paste content, and publish."""
        driver.get("https://www.linkedin.com/feed/")
        start_post_button = self._wait_for_any(
            wait,
            [
                (By.XPATH, "//button[contains(., 'Start a post')]"),
                (By.XPATH, "//button[contains(@aria-label, 'Start a post')]"),
            ],
        )
        start_post_button.click()

        editor = self._wait_for_any(
            wait,
            [
                (By.XPATH, "//div[@role='textbox']"),
                (By.CSS_SELECTOR, "div[contenteditable='true']"),
            ],
        )
        editor.click()

        actions = ActionChains(driver)
        actions.click(editor).key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL).send_keys(post_text).perform()

        post_button = self._wait_for_any(
            wait,
            [
                (By.XPATH, "//button[.//span[text()='Post']]"),
                (By.XPATH, "//button[contains(@aria-label, 'Post')]"),
                (By.XPATH, "//button[contains(., 'Post')]"),
            ],
        )
        if not post_button.is_enabled():
            raise RuntimeError("The LinkedIn Post button is disabled. The editor may not have accepted the text.")

        post_button.click()
        time.sleep(5)
        try:
            wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[@role='dialog']")))
        except TimeoutException:
            self.logger.info("LinkedIn compose dialog did not disappear, but no Selenium exception occurred.")

    @staticmethod
    def _wait_for_any(wait: WebDriverWait, locators: list[tuple[str, str]]):
        """Try multiple selectors because LinkedIn changes its DOM regularly."""
        last_error: Exception | None = None
        for locator in locators:
            try:
                return wait.until(EC.element_to_be_clickable(locator))
            except Exception as exc:
                last_error = exc
        raise last_error or TimeoutException("Could not find a matching LinkedIn element.")


class LinkedInPostingManager:
    """Try the preferred posting mode, then optionally fall back."""

    def __init__(self, settings: Settings) -> None:
        self.settings = settings
        self.logger = get_logger("linkedin_posting")
        self.api_poster = LinkedInAPIPoster(settings)
        self.selenium_poster = LinkedInSeleniumPoster(settings)

    def post(self, post_text: str, preferred_mode: str | None = None, dry_run: bool = False) -> PostingResult:
        """Publish the content using API, Selenium, or both in a fallback chain."""
        if dry_run:
            return PostingResult(True, "dry-run", "Dry run enabled. The post was generated but not sent.")

        mode = (preferred_mode or self.settings.posting_mode or "auto").lower()
        if mode == "auto":
            modes = ["api", "selenium"]
        elif mode in {"api", "selenium"}:
            modes = [mode]
        else:
            raise ValueError("Posting mode must be one of: auto, api, selenium.")

        failures: list[str] = []
        for current_mode in modes:
            result = self.api_poster.post_text(post_text) if current_mode == "api" else self.selenium_poster.post_text(post_text)
            if result.success:
                return result
            failures.append(f"{current_mode}: {result.message}")
            self.logger.warning("Posting via %s failed: %s", current_mode, result.message)

        return PostingResult(False, mode, " | ".join(failures))
