from base.action import BaseO365Action


class CheckInboxAction(BaseO365Action):
    def run(self, limit=10, unread_only=False):
        messages = []

        if self.account.is_authenticated:
            self.logger.debug(self.account.get_current_user())

            mailbox = self.account.mailbox()
            inbox = mailbox.inbox_folder()

            for message in inbox.get_messages(
                limit=limit, download_attachments=False
            ):
                if unread_only and message.is_read:
                    continue

                sender = message.sender.address if message.sender else None

                messages.append(
                    {
                        "id": message.object_id,
                        "subject": message.subject,
                        "sender_email_address": sender,
                        "is_read": message.is_read,
                        "received": str(message.received) if message.received else None,
                        "has_attachments": message.has_attachments,
                        "importance": message.importance,
                    }
                )
        else:
            self.logger.error("Not Authenticated, Check Config and/or generate a token")
            self.logger.debug(f"Authentication Info:\n{self.account.con.__dict__}")

        return messages
