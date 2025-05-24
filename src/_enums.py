from enum import Enum


class OlAutoDiscoverConnectionMode(Enum):
    '''
    Specifies the type of connection to the Exchange server for the
    auto-discovery service.

    - `UNKNOWN: 0` - Other/unknown connection or no connection.
    - `EXTERNAL: 1` -  Connection is over the Internet.
    - `INTERNAL: 2` -  Connection is over the Intranet.
    - `INTERNAL_DOMAIN: 3` -  Connection is same domain over the Intranet.
    '''
    UNKNOWN         = 0
    EXTERNAL        = 1
    INTERNAL        = 2
    INTERNAL_DOMAIN = 3


class OlAccountType(Enum):
    '''
    Specifies the type of an Account.
    
    - `EXCHANGE: 0` - An Exchange account.
    - `IMAP: 1` - An IMAP account.
    - `POP3: 2` - A POP3 account.
    - `HTTP: 3` - An HTTP account.
    - `EAS: 4` - An account using Exchange ActiveSync (EAS).
    - `OTHER_ACCOUNT: 5` - Other/unknown account.
    '''
    EXCHANGE       = 0
    IMAP           = 1
    POP3           = 2
    HTTP           = 3
    EAS            = 4
    OTHER_ACCOUNT  = 5


class OlExchangeConnectionMode(Enum):
    '''
    Specifies whether the account is connected to an Exchange server and if so,
    the connection mode.
    
    - `NO_EXCHANGE: 0` - The account does not use an Exchange server.
    - `OFFLINE: 100` - The account is not connected to an Exchange server and
       is in the classic offline mode. This also occurs when the user selects
       Work Offline from the File menu.
    - `CACHED_OFFLINE: 200` - The account is using cached Exchange mode and the
      user has selected Work Offline from the File menu.
    - `DISCONNECTED: 300` - The account has a disconnected connection to the
      Exchange server.
    - `CACHED_DISCONNECTED: 400` - The account is using cached Exchange mode
      with a disconnected connection to the Exchange server.
    - `CACHED_CONNECTED_HEADERS: 500` - The account is using cached Exchange
      mode on a dial-up or slow connection with the Exchange server, such that
      only headers are downloaded. Full item bodies and attachments remain on
      the server. The user can also select this state manually regardless of
      connection speed.
    - `CACHED_CONNECTED_DRIZZLE: 600` - The account is using cached Exchange
      mode such that headers are downloaded first, followed by the bodies and
      attachments of full items.
    - `CACHED_CONNECTED_FULL: 700` - The account is using cached Exchange mode
      on a Local Area Network or a fast connection with the Exchange server.
      The user can also select this state manually, disabling auto-detect logic
      and always downloading full items regardless of connection speed.
    - `ONLINE: 800` - The account is connected to an Exchange server and is in
      the classic online mode.
    '''
    NO_EXCHANGE               = 0
    OFFLINE                   = 100
    CACHED_OFFLINE            = 200
    DISCONNECTED              = 300
    CACHED_DISCONNECTED       = 400
    CACHED_CONNECTED_HEADERS  = 500
    CACHED_CONNECTED_DRIZZLE  = 600
    CACHED_CONNECTED_FULL     = 700
    ONLINE                    = 800

