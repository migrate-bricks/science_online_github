import os
from typing import Optional
from packaging import version
from webdriver_manager.core.logger import log

from webdriver_manager.core.download_manager import DownloadManager
from webdriver_manager.core.driver_cache import DriverCacheManager
from webdriver_manager.core.manager import DriverManager
from webdriver_manager.core.os_manager import OperationSystemManager, ChromeType
from webdriver_manager.drivers.chrome import ChromeDriver


class ExtChromeDriver(ChromeDriver):
    def get_driver_download_url(self, os_type):
        driver_version_to_download = self.get_driver_version_to_download()
        # For Mac ARM CPUs after version 106.0.5249.61 the format of OS type changed
        # to more unified "mac_arm64". For newer versions, it'll be "mac_arm64"
        # by default, for lower versions we replace "mac_arm64" to old format - "mac64_m1".
        if version.parse(driver_version_to_download) < version.parse("106.0.5249.61"):
            os_type = os_type.replace("mac_arm64", "mac64_m1")

        if version.parse(driver_version_to_download) >= version.parse("115"):
            if os_type == "mac64":
                os_type = "mac-x64"
            if os_type in ["mac_64", "mac64_m1", "mac_arm64"]:
                os_type = "mac-arm64"

            modern_version_url = self.get_url_for_version_and_platform(driver_version_to_download, os_type)
            log(f"Modern chrome version {modern_version_url}")
            return modern_version_url

        return f"{self._url}/{driver_version_to_download}/{self.get_name()}_{os_type}.zip"

    def get_browser_type(self):
        return self._browser_type

    def get_latest_release_version(self):
        determined_browser_version = self.get_browser_version_from_os()
        log(f"Get LATEST {self._name} version for {self._browser_type}")
        if determined_browser_version is not None and version.parse(determined_browser_version) >= version.parse("115"):
            url = "https://registry.npmmirror.com/-/binary/chrome-for-testing"
            response = self._http_client.get(url)
            response_list = response.json()
            determined_browser_version = self.get_version_form_net(determined_browser_version, response_list)
            if determined_browser_version.endswith("/"):
                determined_browser_version = determined_browser_version[:-1]
            return determined_browser_version
            # Remove the build version (the last segment) from determined_browser_version for version < 113
        determined_browser_version = ".".join(determined_browser_version.split(".")[:3])
        latest_release_url = (
            self._latest_release_url
            if (determined_browser_version is None)
            else f"{self._latest_release_url}_{determined_browser_version}"
        )
        resp = self._http_client.get(url=latest_release_url)
        return resp.text.rstrip()

    def get_version_form_net(self, os_version, net_versions):
        for v in net_versions:
            if os_version in v["name"]:
                return v["name"]
        raise Exception(f"No such driver version {os_version} for {self._browser_type}")

    def get_url_for_version_and_platform(self, browser_version, platform):
        base_url = f"https://registry.npmmirror.com/-/binary/chrome-for-testing/{browser_version}/"

        platform_path_map = {
            'linux64': 'linux64/chromedriver-linux64.zip',
            'mac-x64': 'mac-x64/chromedriver-mac-x64.zip',
            'mac-arm64': 'mac-arm64/chromedriver-mac-arm64.zip',
            'win32': 'win32/chromedriver-win32.zip',
            'win64': 'win64/chromedriver-win64.zip',
        }

        download_url = base_url + platform_path_map[platform]
        return download_url


class ChromeDriverManager(DriverManager):
    def __init__(
        self,
        driver_version: Optional[str] = None,
        name: str = "chromedriver",
        url: str = "https://registry.npmmirror.com/-/binary/chromedriver",
        latest_release_url: str = "https://registry.npmmirror.com/-/binary/chromedriver/LATEST_RELEASE",
        chrome_type: str = ChromeType.GOOGLE,
        download_manager: Optional[DownloadManager] = None,
        cache_manager: Optional[DriverCacheManager] = None,
        os_system_manager: Optional[OperationSystemManager] = None

    ):
        super().__init__(
            download_manager=download_manager,
            cache_manager=cache_manager,
            os_system_manager=os_system_manager
        )

        self.driver = ExtChromeDriver(
            name=name,
            driver_version=driver_version,
            url=url,
            latest_release_url=latest_release_url,
            chrome_type=chrome_type,
            http_client=self.http_client,
            os_system_manager=os_system_manager
        )

    def install(self) -> str:
        driver_path = self._get_driver_binary_path(self.driver)
        os.chmod(driver_path, 0o755)
        return driver_path