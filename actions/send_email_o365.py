from base.action import BaseO365Action


class SendEmailAction(BaseO365Action):
    def run(self, subject, body, to_recipients, store):
        mail = None
        if self.account.is_authenticated:
            self.logger.debug(self.account.get_current_user())
            mail = self.account.new_message()
            mail.to.add(to_recipients)
            mail.subject = subject
            mail.body = body
            mail.send(save_to_sent_folder=store)
        else:
            self.logger.error("Not Authenticated, Check Config and/or generate a token")
            self.logger.debug(f"Authentication Info:\n{self.account.con.__dict__}")
        return mail
