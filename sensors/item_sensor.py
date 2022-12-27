import time
from datetime import datetime, timedelta

import pytz
from exchangelib import (DELEGATE, Account, Configuration, EWSDateTime,
                         EWSTimeZone, ServiceAccount)
from O365 import Account, FileSystemTokenBackend, Message, MSGraphProtocol
from st2reactor.sensor.base import PollingSensor

TIME_FORMAT = '%Y-%m-%dT%H:%M:%S'
LOOKBACK = 100  # number of seconds to look for past emails
TN = 'O365OAuthToken'  # name of the token file


class ItemSensor(PollingSensor):
    def __init__(self, sensor_service, config=None, poll_interval=None):
        super(ItemSensor, self).__init__(sensor_service=sensor_service, config=config,
                                         poll_interval=poll_interval)
        # self._logger = self.sensor_service.get_logger(name=self.__class__.__name__)
        # self._stop = False
        # self._store_key = 'exchange.item_sensor_date_str'
        # self._timezone = EWSTimeZone.timezone(config['timezone'])
        # self._credentials = ServiceAccount(
        #     username=config['username'],
        #     password=config['password'])
        # self.primary_smtp_address = config['primary_smtp_address']
        # self.sensor_folder = config['sensor_folder']
        # try:
        #     self.server = config['server']
        #     self.autodiscover = False if self.server is not None else True
        # except KeyError:
        #     self.autodiscover = True

        self._logger = self._sensor_service.get_logger(name=self.__class__.__name__)
        self._stop = False
        # self._store_key = 'exchange.item_sensor_date_str'
        # self._timezone = EWSTimeZone.from_ms_id(config['timezone'])
        self.sensor_folder = config['sensor_folder'] or 'Inbox'
        # self.user = self.config["user"]
        self.tenant_id = self.config['tenant_id']
        self.scopes = self.config['scopes']
        self.protocol = MSGraphProtocol()
        self.credentials = (self.config['client_id'], self.config['client_secret'])
        self.token_backend = FileSystemTokenBackend(
            token_path='/etc/st2/tokens', token_filename=TN)
        # self.token_backend = StackstormTokenBackend(self.user.split('@')[0], self._sensor_service)
        self.account = Account(self.credentials, auth_flow_type='authorization', token_backend=self.token_backend,
                               tenant_id=self.tenant_id, protocol=self.protocol)

    def setup(self):
        # if self.autodiscover:
        #     self.account = Account(
        #         primary_smtp_address=self.primary_smtp_address,
        #         credentials=self._credentials,
        #         autodiscover=True,
        #         access_type=DELEGATE)
        # else:
        #     ms_config = Configuration(
        #         server=self.server,
        #         credentials=self._credentials)
        #     self.account = Account(
        #         primary_smtp_address=self.primary_smtp_address,
        #         config=ms_config,
        #         autodiscover=False,
        #         access_type=DELEGATE)
        if self.account.con.token_backend.token_path.exists():
            self._logger.info(
                f"Authenticated Token Found")
            self.account.connection.refresh_token()
        else:
            self._logger.warning(
                f"Authenticated Token NOT Found")
        if self.account.is_authenticated:
            self._logger.info(f"Authenticated User: {self.account.get_current_user()}")
        else:
            self._logger.warning("Not Authenticated, Check Config and generate a token")
            self._logger.debug(self.account.__dict__)
        if self.sensor_folder:
            self._logger.info(f"Monitored Folder: {self.sensor_folder}")
        else:
            self._logger.warning("No Folder set for Monitoring")

    def poll(self):
        # stored_date = self._get_last_date()
        # self._logger.info("Stored Date: {}".format(stored_date))
        # if not stored_date:
        #     stored_date = datetime.now()
        # # pylint: disable=no-member
        # start_date = self._timezone.localize(EWSDateTime.from_datetime(stored_date))
        # target = self.account.root.get_folder_by_name(self.sensor_folder)
        # items = target.filter(is_read=False).filter(datetime_received__gt=start_date)

        # self._logger.info("Found {0} items".format(items.count()))

        # for newitem in items:
        #     self._logger.info("Sending trigger for item '{0}'.".format(newitem.subject))
        #     self._dispatch_trigger_for_new_item(newitem=newitem)
        #     self._set_last_date(newitem.datetime_received)
        #     self._logger.info("Updating read status on item '{0}'.".format(newitem.subject))
        #     newitem.is_read = True
        #     newitem.save()

        if self.account.is_authenticated:
            start_date = datetime.utcnow().replace(tzinfo=pytz.utc) - timedelta(LOOKBACK)
            # start_date = self._timezone.localize(EWSDateTime.from_datetime(stored_date))
            self._logger.debug("Selecting Folder {0}".format(self.sensor_folder))
            _folder = self.account.mailbox().get_folder(folder_name=self.sensor_folder)
            # _folder = self.account.root.get_folder_by_name(self.sensor_folder)
            _query = _folder.q()
            # only look back to start_date
            _query = _query.chain('and').on_attribute('received_date_time').greater_equal(
                start_date)
            # only get unread messages
            _query = _query.chain('and').on_attribute('isRead').equals(False)

            _messages = _folder.get_messages(limit=25, query=_query, download_attachments=False)
            # _cnt = 0
            for _message in _messages:

                self._logger.info("Sending trigger for item '{0}'.".format(_message.subject))
                self._dispatch_trigger_for_new_item(newitem=_message)
        else:
            self._logger.warning("Not Authenticated, Check Config and/or generate a token")

    def cleanup(self):
        # This is called when the st2 system goes down. You can perform cleanup operations like
        # closing the connections to external system here.
        pass

    def add_trigger(self, trigger):
        # This method is called when trigger is created
        pass

    def update_trigger(self, trigger):
        # This method is called when trigger is updated
        pass

    def remove_trigger(self, trigger):
        # This method is called when trigger is deleted
        pass

    def _get_last_date(self):
        self._last_date = self._sensor_service.get_value(name=self._store_key)
        if self._last_date is None:
            return None
        return datetime.strptime(self._last_date, TIME_FORMAT)

    def _set_last_date(self, last_date):
        # Check if the last_date value is an EWSDateTime object
        if isinstance(last_date, EWSDateTime):
            self._last_date = last_date.strftime(TIME_FORMAT)
        else:
            self._last_date = time.strftime(TIME_FORMAT, last_date)
        self._sensor_service.set_value(name=self._store_key,
                                       value=self._last_date)

    def _dispatch_trigger_for_new_item(self, newitem: Message):
        # trigger = 'msexchange.exchange_new_item'
        # if isinstance(newitem.datetime_received, EWSDateTime):
        #     datetime_received = newitem.datetime_received.strftime(TIME_FORMAT)
        # else:
        #     datetime_received = str(newitem.datetime_received)

        # payload = {
        #     'item_id': str(newitem.item_id),
        #     "change_key": str(newitem.changekey),
        #     'subject': str(newitem.subject),
        #     'body': str(newitem.body),
        #     'datetime_received': datetime_received,
        # }
        # self._sensor_service.dispatch(trigger=trigger, payload=payload)

        trigger = 'msexchange.exchange_new_item'
        if isinstance(newitem.received, EWSDateTime):
            datetime_received = newitem.received.strftime(TIME_FORMAT)
        else:
            datetime_received = str(newitem.received)

        payload = {
            'item_id': str(newitem.object_id),
            'subject': str(newitem.subject),
            'body': str(newitem.get_body_text()),
            'datetime_received': datetime_received,
        }
        self._logger.info("Sender: '{0}'.".format(newitem.sender))
        self._sensor_service.dispatch(trigger=trigger, payload=payload)
        self._logger.debug("Updating read status on item '{0}'.".format(newitem.subject))
        if newitem.mark_as_read():
            self._logger.info("Marked item as read")
        else:
            self._logger.error(
                "Unable to mark message as Read: '{0}'.".format(newitem.object_id))
