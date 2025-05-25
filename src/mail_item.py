from __future__ import annotations
from typing import TYPE_CHECKING
from win32com.client import CDispatch
from . import _enums

if TYPE_CHECKING:
#     from .account import Account
    from .application import Application
    from .folder import Folder
#     from .namespace import NameSpace


class MailItem:
    '''
    description

    Properties
    ----------

    Methods
    -------

    '''
    
    def __init__(self, folder: Folder, mail_item: CDispatch) -> None:
        self.folder = folder
        self._mail_item = mail_item
        return
    
    @property
    def actions(self) -> CDispatch:
        '''
        Returns an `Actions` collection that represents all the available
        actions for the item. Read-only.
        '''
        return self._mail_item.Actions
    
    @property
    def alternate_recipient_allowed(self) -> bool:
        '''
        Returns `True` if the mail message can be forwarded. Read/write
        '''
        return self._mail_item.AlternateRecipientAllowed
    
    @alternate_recipient_allowed.setter
    def alternate_recipient_allowed(self, value: bool) -> None:
        self._mail_item.AlternateRecipientAllowed = value
        return
    
    @property
    def application(self) -> Application:
        '''
        Returns an `Application` object that represents the parent Outlook
        application for the object. Read-only.
        '''
        return self.folder.application
    
    @property
    def attachments(self) -> CDispatch:
        '''
        Returns an `Attachments` object that represents all the attachments for
        the specified item. Read-only.
        '''
        return self._mail_item.Attachments
    
    @property
    def auto_forwarded(self) -> bool:
        '''
        A Boolean value that returns `True` if the item was automatically
        forwarded. Read/write.
        '''
        return self._mail_item.AutoForwarded
    
    @auto_forwarded.setter
    def auto_forwarded(self, value: bool) -> None:
        self._mail_item.AutoForwarded = value
        return
    
    @property
    def auto_resolved_winner(self) -> bool:
        '''
        Returns a Boolean that determines if the item is a winner of an
        automatic conflict resolution. Read-only.

        Remarks
        -------
        A value of `False` does not necessarily indicate that the item is a
        loser of an automatic conflict resolution. The item could be in
        conflict with another item.

        If an item has `Conflicts.Count` of its `MailItem.conflicts` property
        greater than zero and if its `auto_resolved_winner` property is `True`,
        it is a winner of an automatic conflict resolution. On the other hand,
        if the item is in conflict and has its `auto_resolved_winner` property
        as `False`, it is a loser in an automatic conflict resolution.
        '''
        return self._mail_item.AutoResolvedWinner
    
    @property
    def bcc(self) -> str:
        '''
        Returns a string representing the display list of blind carbon copy
        (BCC) names for a `MailItem`. Read/write.
        '''
        return self._mail_item.BCC
    
    @bcc.setter
    def bcc(self, value: str) -> None:
        self._mail_item.BCC = value
        return
    
    @property
    def billing_information(self) -> str:
        '''
        Returns or sets a string representing the billing information
        associated with the Outlook item. Read/write.

        Remarks
        -------
        This is a free-form text field.
        '''
        return self._mail_item.BillingInformation
    
    @billing_information.setter
    def billing_information(self, value: str) -> None:
        self._mail_item.BillingInformation = value
        return
    
    @property
    def body(self) -> str:
        '''
        Returns or sets a string representing the clear-text body of the
        Outlook item. Read/write.

        Remarks
        -------
        The `MailItem.body_format` property allows you to programmatically
        change the editor that is used for the body of an item.
        '''
        return self._mail_item.Body
    
    @body.setter
    def body(self, value: str) -> None:
        self._mail_item.Body = value
        return
    
    @property
    def body_format(self) -> _enums.OlBodyFormat:
        '''
        Returns or sets an `OlBodyFormat` constant indicating the format of the
        body text. Read/write.

        Remarks
        -------
        The body text format determines the standard used to display the text
        of the message. Microsoft Outlook provides three body text format
        options: Plain Text, Rich Text (RTF), and HTML.

        All text formatting will be lost when the `body_format` property is
        switched from RTF to HTML and vice-versa.
        '''
        body_format_value = self._mail_item.BodyFormat
        return _enums.OlBodyFormat(body_format_value)
    
    @body_format.setter
    def body_format(self, value: int | _enums.OlBodyFormat) -> None:
        if isinstance(value, _enums.OlBodyFormat):
            value = value.value
        else:
            assert value in _enums.OlBodyFormat, \
                f'Invalid value {value} for OlBodyFormat enum.'
        self._mail_item.BodyFormat = value
        return
    
    @property
    def categories(self) -> str:
        '''
        Returns or sets a string representing the categories assigned to the
        Outlook item. Read/write.

        Remarks
        -------
        `categories` is a comma-delimited string of category names that have
        been assigned to an Outlook item.
        '''
        return self._mail_item.Categories
    
    @categories.setter
    def categories(self, value: str) -> None:
        self.Categories = value
        return
    
    @property
    def cc(self) -> str:
        '''
        Returns a string representing the display list of carbon copy (CC)
        names for a `MailItem`. Read/write.

        Remarks
        -------
        This property contains the display names only. The `Recipients`
        collection should be used to modify the CC recipients.
        '''
        return self._mail_item.CC
    
    @cc.setter
    def cc(self, value: str) -> None:
        self._mail_item.CC = value
        return
    
    @property
    def companies(self) -> str:
        '''
        Returns or sets a string representing the names of the companies
        associated with the Outlook item. Read/write.

        Remarks
        -------
        This is a free-form text field.
        '''
        return self._mail_item.Companies
    
    @companies.setter
    def companies(self, value: str) -> None:
        self._mail_item.Companies = value
        return


'''
Potential Attributes:
---------------------
Conflicts
ConversationID
ConversationIndex
ConversationTopic
Copy
CreationTime
DeferredDeliveryTime
Delete
DeleteAfterSubmit
Display
DownloadState
EnableSharedAttachments
EntryID
ExpiryTime
FlagDueBy
FlagIcon
FlagRequest
FlagStatus
FormDescription
Forward
GetConversation
GetIDsOfNames
GetInspector
GetTypeInfo
GetTypeInfoCount
HTMLBody
HasCoverSheet
Importance
InternetCodepage
Invoke
IsConflict
IsIPFax
IsMarkedAsTask
ItemProperties
LastModificationTime
Links
MAPIOBJECT
MarkAsTask
MarkForDownload
MessageClass
Mileage
Move
NoAging
OriginatorDeliveryReportRequested
OutlookInternalVersion
OutlookVersion
Parent
Permission
PermissionService
PermissionTemplateGuid
PrintOut
PropertyAccessor
QueryInterface
RTFBody
ReadReceiptRequested
ReceivedByEntryID
ReceivedByName
ReceivedOnBehalfOfEntryID
ReceivedOnBehalfOfName
ReceivedTime
RecipientReassignmentProhibited
Recipients
Release
ReminderOverrideDefault
ReminderPlaySound
ReminderSet
ReminderSoundFile
ReminderTime
RemoteStatus
Reply
ReplyAll
ReplyRecipientNames
ReplyRecipients
RetentionExpirationDate
RetentionPolicyName
Save
SaveAs
SaveSentMessageFolder
Saved
Send
SendUsingAccount
Sender
SenderEmailAddress
SenderEmailType
SenderName
Sensitivity
Sent
SentOn
SentOnBehalfOfName
Session
ShowCategoriesDialog
Size
Subject
Submitted
TaskCompletedDate
TaskDueDate
TaskStartDate
TaskSubject
To
ToDoTaskOrdinal
UnRead
UserProperties
VotingOptions
VotingResponse



Methods
-------
AddBusinessCard
ClearConversationIndex
ClearTaskFlag
Close

'''