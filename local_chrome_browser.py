from playwright.sync_api import sync_playwright, Page
import time
import os
import random
from typing import Optional, List, Dict

class LocalChromeBrowser:
    def __init__(self, origin: Optional[str] = None, 
                 pfxPath: Optional[str] = None, 
                 passphrase: Optional[str] = None,
                 cert_config: Optional[List[Dict]] = None):
        """
        Initialize browser with optional certificate parameters.
        All certificate parameters must be provided together or none at all.
        
        Args:

        origin: Union[str, None] (optional) Exact origin that the certificate is valid for. Origin includes https protocol, a hostname and optionally a port.

        pfxPath: Union[str, pathlib.Path] (optional) Path to the PFX or PKCS12 encoded private key and certificate chain.
 and certificate chain.

        passphrase: str (optional) Passphrase for the private key (PEM or PFX).
        cert_config: List[Dict] (optional) Configuration for client certificates, if any.
        If provided, it should be a list of dictionaries with keys 'origin', 'pfxPath', and 'passphrase'.
        """
        self.playwright = None
        self.context = None
        self.page: Optional[Page] = None
        
        # Validate certificate parameters
        cert_params = [origin, pfxPath, passphrase]
        if any(cert_params) and not all(cert_params):
            raise ValueError("All certificate parameters (origin, path, and password) must be provided together")
            
        self.cert_config = cert_config or None
        if all(cert_params):
            self.cert_config = [{
                "origin": origin,
                "pfxPath": pfxPath,
                "passphrase": passphrase
            }]

    def __enter__(self):
        # Initialize playwright
        self.playwright = sync_playwright().start()
        
        # Get user profile directory
        user_profile = os.path.join(
            os.environ['LOCALAPPDATA'],
            'Google/Chrome/User Data/Default'
        )
        
        # Prepare context options
        context_options = {
            'user_data_dir': user_profile,
            'channel': "chrome",
            'headless': False,
            'executable_path': "C:/Program Files/Google/Chrome/Application/chrome.exe",
            'ignore_https_errors': True,
            'args': [
                # Basic automation flags
                '--disable-blink-features=AutomationControlled',
                '--disable-automation',
                '--disable-extensions',
                '--disable-infobars',
                '--no-sandbox',
                '--disable-dev-shm-usage',
                '--window-size=1920,1080',
                f'--window-position={random.randint(0,100)},{random.randint(0,100)}',
                
                # Session restore prevention
                '--no-first-run',
                '--no-default-browser-check',
                '--disable-session-crashed-bubble',
                '--disable-restore-session-state',
                '--disable-sync',
                '--disable-crash-reporter',
                # '--incognito',  # Use incognito mode to avoid session restore
                '--start-maximized',
                
                # Additional session handling
                # '--disable-features=TranslateUI,ChromeWhatsNew',
                # '--disable-popup-blocking',
                # '--disable-notifications',
                # '--disable-save-password-bubble',
                # '--disable-component-update',
                # '--noerrdialogs',
                # '--silent-launch',
                # '--disable-background-mode',
                # '--disable-breakpad',
                # '--disable-default-apps',
                # '--no-startup-window',
                # '--test-type',
                # '--enable-precise-memory-info',
                '--force-empty-session-state',  # Force empty session state
            ],
            'bypass_csp': True,
            'viewport': {'width': 1920, 'height': 1080},
        }

        # Add certificates if configured
        if self.cert_config:
            context_options['client_certificates'] = self.cert_config
        
        # Launch persistent context
        self.context = self.playwright.chromium.launch_persistent_context(**context_options)
        
        # Create page and add stealth scripts
        self.page = self.context.new_page()
        self._add_stealth_scripts()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.context:
            self.context.close()
        if self.playwright:
            self.playwright.stop()

    def _add_stealth_scripts(self):
        """Add anti-detection scripts to the page"""
        self.page.add_init_script("""
            // Overwrite the automation-related properties
            const overwriteProperties = {
                webdriver: undefined,
                __driver_evaluate: undefined,
                __webdriver_evaluate: undefined,
                __selenium_evaluate: undefined,
                __fxdriver_evaluate: undefined,
                __driver_unwrapped: undefined,
                __webdriver_unwrapped: undefined,
                __selenium_unwrapped: undefined,
                __fxdriver_unwrapped: undefined,
                _Selenium_IDE_Recorder: undefined,
                calledSelenium: undefined,
                _selenium: undefined,
                callSelenium: undefined,
                _WEBDRIVER_ELEM_CACHE: undefined,
                ChromeDriverw: undefined,
                domAutomation: undefined,
                domAutomationController: undefined,
            };

            Object.keys(overwriteProperties).forEach(prop => {
                Object.defineProperty(window, prop, {
                    get: () => overwriteProperties[prop],
                    set: () => {}
                });
                Object.defineProperty(navigator, prop, {
                    get: () => overwriteProperties[prop],
                    set: () => {}
                });
            });

            delete navigator.__proto__.webdriver;
        """)

    def goto(self, url: str):
        """Navigate to a URL"""
        self.page.goto(url, wait_until="networkidle")
        return self

    def random_delay(self, min_seconds: float = 1, max_seconds: float = 3):
        """Add random delay between actions"""
        time.sleep(random.uniform(min_seconds, max_seconds))
        return self

# Example usage:
if __name__ == "__main__":
    with LocalChromeBrowser() as browser:
        browser.goto("https://enel.service-now.com/navpage.do")
        browser.random_delay()
        # Access the page directly for Playwright Page methods
        browser.page.click("text=Comparativas")
        browser.random_delay()
        browser.page.click("text=Panel de información")

    # Basic usage without certificates
    with LocalChromeBrowser() as browser:
        browser.goto("https://example.com")

    # Usage with certificates
    with LocalChromeBrowser(
        origin="https://your-server.com",
        pfxPath="./path/to/cert.pfx",
        passphrase="your-password"
    ) as browser:
        browser.goto("https://your-server.com")
        # Your code here...