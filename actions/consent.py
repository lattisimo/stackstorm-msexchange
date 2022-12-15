from lib.action import BaseO365Action
import webbrowser


class UserConsent(BaseO365Action):
    def run(self):
        consent_url, _ = self.consent_yo()
        webbrowser.open_new_tab(consent_url)
        return (True, consent_url)
