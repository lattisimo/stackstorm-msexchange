from base.action import BaseO365Action


class GetToken(BaseO365Action):
    def run(self, consent_url):
        return (True, self.token_yo(consent_url))
