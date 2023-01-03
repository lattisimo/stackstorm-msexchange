from base.action import BaseO365Action
# import codecs


class SendEmailAction(BaseO365Action):
    def run(self, subject, body, to_recipients, store, body_type='text'):
        mail = None
        if self.account.is_authenticated:
            self.logger.debug(self.account.get_current_user())
            mail = self.account.new_message()
            mail.to.add(to_recipients)
            mail.subject = subject
            mail.body_type = body_type
            mail.body = body  # codecs.decode(codecs.encode(body, "utf-8"))
            mail.send(save_to_sent_folder=store)
        else:
            self.logger.error("Not Authenticated, Check Config and/or generate a token")
            self.logger.debug(f"Authentication Info:\n{self.account.con.__dict__}")
        return mail