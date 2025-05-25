from __future__ import annotations
from typing import Iterator, TYPE_CHECKING
from collections import UserList
from win32com.client import CDispatch
from . import _enums

if TYPE_CHECKING:
    from .account import Account
    from .application import Application
    from .namespace import NameSpace



class Folder(UserList[CDispatch]):
    '''
    Represents an Outlook folder.

    Properties
    ----------
    address_book_name : str
        Returns or sets a string that indicates the Address Book name for
        the Folder object representing a Contacts folder. Read/write.
    application : Application
        Returns an Application object that represents the parent Outlook
        application for the object.
    current_view : CDispatch
        Returns a `View` object representing the current view. Read-only.
    custom_views_only : bool
        Returns or sets a bool that determines which views are displayed on
        the `View` menu for a given folder. Read/write.
    default_item_type : OlItemType
        Returns a constant from the `OlItemType` enumeration indicating the
        default Outlook item type contained in the folder. Read-only.
    default_message_class : str
        Returns a string representing the default message class for items in
    description : str
        Returns or sets a string representing the description of the folder.
        Read/write.
    entry_id : str
        Returns a string representing the unique Entry ID of the object.
        Read-only.
    folder_path : str
        Returns a string that indicates the path of the current folder.
        Read-only.
    folders : list[Folder]
        Returns a list of `Folder` objects that represent all the sub-folders
        contained in the specified folder. Read-only.
    in_app_folder_sync_object : bool
        Returns or sets a Boolean that determines if the specified folder will
        be synchronized with the e-mail server. Read/write.
    is_share_point_folder : bool
        Returns a Boolean that determines if the folder is a Microsoft
        SharePoint Foundation folder. Read-only.
    name : str
        Returns or sets a string value that represents the display name for the
        object. Read/write.
    property_accessor : CDispatch
        Returns a `PropertyAccessor` object that supports creating, getting,
        setting, and deleting properties of the parent Folder object.
        Read-only.
    session : NameSpace
        Returns the `NameSpace` object for the current session.
    show_as_outlook_ab : bool
        Returns or sets a Boolean value that specifies whether the contact
        items folder will be displayed as an address list in the Outlook
        Address Book. Read/write.
    show_item_count : OlShowItemCount
        Sets or returns a constant in the OlShowItemCount enumeration that
        indicates whether to display the number of unread messages in the
        folder or the total number of items in the folder in the Navigation
        Pane. Read/write.
    store : Account
        Returns an `Account` object representing the store that contains the
        `Folder` object. Read-only.
    store_id : str
        Returns a string indicating the store ID for the folder. Read-only.
    unread_item_count : int
        Returns an integer value indicating the number of unread items in the
        folder. Read-only.
    user_defined_properties : CDispatch
        Returns a `UserDefinedProperties` object that represents the
        user-defined custom properties for the `Folder` object. Read-only.
    views : CDispatch
        Returns the Views collection of the Folder object. Read-only.
    web_view_on : bool
        Returns or sets a bool indicating the Web view state for a folder.
        Read/write.
    web_view_url : str
        Returns or sets a string indicating the URL of the Web page that is
        assigned to a folder. Read/write. 

    Methods
    -------
    add_to_pffavorites()
        Adds a Microsoft Exchange public folder to the public folder's
        Favorites folder.
    copy_to(destination_folder)
        Copies the current folder in its entirety to the destination folder.
    delete()
        Deletes an object from the collection.
    display()
        Displays a new `Explorer` object for the folder.
    get_calendar_exporter()
        Creates a `CalendarSharing` object for the specified Folder.
    get_custom_icon()
        Returns an `StdPicture` object that represents the custom icon for the
        folder.
    get_explorer(display_mode=0)
        Returns an `Explorer` object that represents a new, inactive `Explorer`
        object initialized with the specified folder as the current folder.
    get_storage(storage_identifier, storage_identifier_type)
        Gets a `StorageItem` object on the parent `Folder` to store data for an
        Outlook solution.
    get_table(filter_=None, table_contents=0)
        Obtains a `Table` object that contains items filtered by `filter_`.
    items()
        Returns a list of `Item` objects from the specified folder. Read-only.
    move_to(destination_folder)
        Moves a folder to the specified destination folder.
    set_custom_icon(picture)
        Sets a custom icon that is specified by `picture` for the folder.

    Remarks
    -------
    This is a .NET interface derived from a COM co-class that is required by
    managed code for interoperability with the corresponding COM object. Use
    this interface to access all method, property, and event members of the COM
    object `Folder`. Refer to this topic for information about the COM object.

    A `Folder` object can contain other `Folder` objects, as well as Outlook
    items. Use the `folders` property of a `NameSpace` object or another
    `Folder` object to return the set of folders in a `NameSpace` or under a
    folder. You can navigate nested folders by starting from a top-level
    folder, say the Inbox, and using a combination of the `folders` property,
    which returns the set of folders underneath a Folder object in the
    hierarchy.

    There is a set of folders within an Outlook data store that supports the
    default functionality of Outlook. Use `get_default_folder`, specifying an
    index that is one of the constants in the `OlDefaultFolders` enumeration to
    return one of the default Outlook folders in the Outlook `NameSpace`
    object.

    While generally it is a good practice to place items that serve the same
    functionality in the same folder, a folder can contain items of different
    types. For example, by default, the Calendar folder can contain
    `AppointmentItem` and `MeetingItem` objects, and the Contacts folder can
    contain `ContactItem` and `DistListItem` objects. In general, when
    enumerating items in a folder, do not assume the type of an item in the
    folder.

    Use the `add` method to add a folder to the `Folders` object. The add
    method has an optional argument that can be used to specify the type of
    items that can be stored in that folder. By default, folders created
    inside another folder inherit the type of the parent folder.

    Note that when items of a specific type are saved, they are saved directly
    into their corresponding default folder. For example, when the
    `get_associated_appointment` method is applied to a `MeetingItem` in the
    Inbox folder, the appointment that is returned will be saved to the default
    Calendar folder.
    '''

    def __init__(self, account: Account, folder: CDispatch) -> None:
        self.account = account
        self._folder = folder
        super().__init__(self._folder.Items)
        return
    
    def __repr__(self) -> str:
        return f"<Folder '{self.folder_path}'>"

    @property
    def address_book_name(self) -> str:
        '''
        Returns or sets a string that indicates the Address Book name for
        the Folder object representing a Contacts folder. Read/write.
        '''
        return self._folder.AddressBookName
    
    @address_book_name.setter
    def address_book_name(self, name: str) -> None:
        self._folder.AddressBookName = name
        return

    @property
    def application(self) -> Application:
        '''
        Returns an Application object that represents the parent Outlook
        application for the object.
        '''
        return self.account.application

    @property
    def current_view(self) -> CDispatch:
        '''
        Returns a `View` object representing the current view. Read-only.
        '''
        return self._folder.CurrentView

    @property
    def custom_views_only(self) -> bool:
        '''
        Returns or sets a bool that determines which views are displayed on
        the `View` menu for a given folder. Read/write.
        '''
        return self._folder.CustomViewsOnly
    
    @custom_views_only.setter
    def custom_views_only(self, value: bool) -> None:
        self._folder.CustomViewsOnly = value
        return

    @property
    def default_item_type(self) -> _enums.OlItemType:
        '''
        Returns a constant from the `OlItemType` enumeration indicating the
        default Outlook item type contained in the folder. Read-only.
        '''
        default_item_type = self._folder.DefaultItemType
        return _enums.OlItemType(default_item_type)

    @property
    def default_message_class(self) -> str:
        '''
        Returns a string representing the default message class for items in
        the folder. Read-only.'''
        return self._folder.DefaultMessageClass

    @property
    def description(self) -> str:
        '''
        Returns or sets a string representing the description of the folder.
        Read/write.
        '''
        return self._folder.Description

    @property
    def entry_id(self) -> str:
        '''
        Returns a string representing the unique Entry ID of the object.
        Read-only.
        '''
        return self._folder.EntryID

    @property
    def folder_path(self) -> str:
        '''
        Returns a string that indicates the path of the current folder.
        Read-only.
        '''
        return self._folder.FolderPath

    @property
    def folders(self) -> list[Folder]:
        '''
        Returns a list of `Folder` objects that represent all the sub-folders
        contained in the specified folder. Read-only.
        '''
        folders = self._folder.Folders
        return [Folder(self.account, f) for f in folders]

    @property
    def in_app_folder_sync_object(self) -> bool:
        '''
        Returns or sets a Boolean that determines if the specified folder will
        be synchronized with the e-mail server. Read/write.
        '''
        return self._folder.InAppFolderSyncObject
    
    @in_app_folder_sync_object.setter
    def in_app_folder_sync_object(self, value: bool) -> None:
        self._folder.InAppFolderSyncObject = value
        return

    @property
    def is_share_point_folder(self) -> bool:
        '''
        Returns a Boolean that determines if the folder is a Microsoft
        SharePoint Foundation folder. Read-only.
        '''
        return self._folder.IsSharePointFolder

    @property
    def name(self) -> str:
        '''
        Returns or sets a string value that represents the display name for the
        object. Read/write.
        '''
        return self._folder.Name
    
    @name.setter
    def name(self, value: str) -> None:
        self._folder.Name = value
        return

    @property
    def property_accessor(self) -> CDispatch:
        '''
        Returns a `PropertyAccessor` object that supports creating, getting,
        setting, and deleting properties of the parent Folder object.
        Read-only.
        '''
        return self._folder.PropertyAccessor

    @property
    def session(self) -> NameSpace:
        '''
        Returns the `NameSpace` object for the current session.
        '''
        return self.account.namespace

    @property
    def show_as_outlook_ab(self) -> bool:
        '''
        Returns or sets a Boolean value that specifies whether the contact
        items folder will be displayed as an address list in the Outlook
        Address Book. Read/write.
        '''
        return self._folder.ShowAsOutlookAB
    
    @show_as_outlook_ab.setter
    def show_as_outlook_ab(self, value: bool) -> None:
        self._folder.ShowAsOutlookAB = value
        return

    @property
    def show_item_count(self) -> _enums.OlShowItemCount:
        '''
        Sets or returns a constant in the OlShowItemCount enumeration that
        indicates whether to display the number of unread messages in the
        folder or the total number of items in the folder in the Navigation
        Pane. Read/write.
        '''
        show_item_count = self._folder.ShowItemCount
        return _enums.OlShowItemCount(show_item_count)
    
    @show_item_count.setter
    def show_item_count(
            self,
            show_item_count: int | _enums.OlShowItemCount
    ) -> None:
        if isinstance(show_item_count, _enums.OlShowItemCount):
            show_item_count = show_item_count.value
        self._folder.ShowItemCount = show_item_count
        return

    @property
    def store(self) -> Account:
        '''
        Returns an `Account` object representing the store that contains the
        `Folder` object. Read-only.
        '''
        return self.account

    @property
    def store_id(self) -> str:
        '''
        Returns a string indicating the store ID for the folder. Read-only.
        '''
        return self._folder.StoreID

    @property
    def unread_item_count(self) -> int:
        '''
        Returns an integer value indicating the number of unread items in the
        folder. Read-only.
        '''
        return self._folder.UnReadItemCount

    @property
    def user_defined_properties(self) -> CDispatch:
        '''
        Returns a `UserDefinedProperties` object that represents the
        user-defined custom properties for the `Folder` object. Read-only.
        '''
        return self._folder.UserDefinedProperties

    @property
    def views(self) -> CDispatch:
        '''
        Returns the Views collection of the Folder object. Read-only.
        '''
        return self._folder.Views

    @property
    def web_view_on(self) -> bool:
        '''
        Returns or sets a bool indicating the Web view state for a folder.
        Read/write.
        '''
        return self._folder.WebViewOn
    
    @web_view_on.setter
    def web_view_on(self, value: bool) -> None:
        self._folder.WebViewOn = value
        return

    @property
    def web_view_url(self) -> str:
        '''
        Returns or sets a string indicating the URL of the Web page that is
        assigned to a folder. Read/write. 
        '''
        return self._folder.WebViewURL
    
    @web_view_url.setter
    def web_view_url(self, value: str) -> None:
        self._folder.WebViewURL = value
        return

    # Methods

    def add_to_pffavorites(self) -> None:
        '''
        Adds a Microsoft Exchange public folder to the public folder's
        Favorites folder.
        '''
        return self._folder.AddToPFFavorites()

    def copy_to(self, destination_folder: Folder) -> Folder:
        '''
        Copies the current folder in its entirety to the destination folder.

        Parameters
        ----------
        destination_folder : Folder
            Required `Folder` object that represents the destination folder.

        Returns
        -------
        Folder
            A `Folder` object that represents the new copy of the current
            folder.

        Remarks
        -------
        Setting the `REG_MULTI_SZ` value, `DisableCrossAccountCopy`, in
        `HKCU\\Software\\Microsoft\\Office\\14.0\\Outlook` in the Windows
        registry has the side effect of disabling this method.
        '''
        _dest_folder = destination_folder._folder
        _new_folder = self._folder.CopyTo(_dest_folder)
        return Folder(self.account, _new_folder)

    def delete(self) -> None:
        '''
        Deletes an object from the collection.
        '''
        return self._folder.Delete()

    def display(self) -> None:
        '''
        Displays a new `Explorer` object for the folder.
        '''
        return self._folder.Display()

    def get_calendar_exporter(self) -> CDispatch:
        '''
        Creates a `CalendarSharing` object for the specified Folder.

        Returns
        -------
        CDispatch
            A `CalendarSharing` object for the specified folder.

        Remarks
        -------
        The `get_calendar_exporter` method automatically sets the defaults for
        the `CalendarSharing` class to the standard default options used by the
        `Folder` object. The `get_calendar_exporter` method can only be used on
        calendar folders. An error occurs if you use the method on `Folder`
        objects that represent other folder types.

        **Note:** The `CalendarSharing` object only supports exporting the
        iCalendar (`.ics`) file format.
        '''
        return self._folder.GetCalendarExporter()

    def get_custom_icon(self) -> CDispatch:
        '''
        Returns an `StdPicture` object that represents the custom icon for the
        folder.

        Returns
        -------
        CDispatch
            A `StdPicture` object that represents a custom icon for the folder.

        Remarks
        -------
        The returned `StdPicture` object has its `Type` property equal to
        `PICTYPE_ICON` or `PICTYPE_BITMAP`.

        `get_custom_icon` returns `None` if the folder does not have a custom
        folder icon.

        You can only call `get_custom_icon` from code that runs in-process as
        Outlook. A `StdPicture` object cannot be marshaled across process
        boundaries. If you attempt to call `get_custom_icon` from
        out-of-process code, an exception is raised.
        '''
        return self._folder.GetCustomIcon()

    def get_explorer(
            self,
            display_mode: int | _enums.OlFolderDisplayMode = 0
    ) -> CDispatch:
        '''
        Returns an `Explorer` object that represents a new, inactive `Explorer`
        object initialized with the specified folder as the current folder.

        Parameters
        ----------
        display_mode
            The display mode of the folder. Can be one of the constants in the
            `OlFolderDisplayMode` enumeration.

        Returns
        -------
        CDispatch
            An `Explorer` object that represents a new, inactive `Explorer`
            initialized with the specified folder as the current folder.

        Remarks
        -------
        This method is useful for returning a new `Explorer` object in which to
        display the folder, as opposed to using the `active_explorer` method
        and setting the `current_folder` property.

        The `display` method can be used to activate or show the Explorer.

        The `get_explorer` method takes an optional argument of an
        `OlFolderDisplayMode` constant.

        By default, the new `Explorer` will be displayed in the Normal mode
        (`NORMAL`) with all interface elements displayed: a message panel on
        the right and the Navigation Pane on the left. The exception to this
        rule is when you are calling `get_explorer` on delegated folders that
        are in No-Navigation mode (`NO_NAVIGATION`) by default. You can apply
        more restrictions to a default mode, but you cannot lessen the
        restrictions by changing the `OlFolderDisplayMode`.

        The explorer can also be displayed in Folder-Only mode (`FOLDER_ONLY`).
        This mode is essentially the same as the Normal mode in that it too
        displays the Navigation Pane on the left.

        The most restrictive mode you can use is No-Navigation mode. In this
        mode, the Explorer will display with no folder list, no drop-down
        folder list, and any "Go"-type menu/command bar options should be
        disabled. Basically, the user should not be able to navigate to any
        other folder within that Explorer window. By default, a delegated
        (shared) folder appears in No-Navigation mode.
        '''
        if isinstance(display_mode, _enums.OlFolderDisplayMode):
            display_mode = display_mode.value
        return self._folder.GetExplorer(display_mode)

    def get_storage(
            self,
            storage_identifier: str,
            storage_identifier_type: int | _enums.OlStorageIdentifierType
    ) -> CDispatch:
        '''
        Gets a `StorageItem` object on the parent `Folder` to store data for an
        Outlook solution.

        Parameters
        ----------
        storage_identifier : str
            An identifier for the `StorageItem` object; depending on the
            identifier type, the value can represent an Entry ID, a message
            class, or a subject.
        storage_identifier_type : int | OlStorageIdentifierType
            Specifies the type of identifier for the `StorageItem` object.
            
        Returns
        -------
        CDispatch
            A `StorageItem` object that is used to store data for a solution.

        Remarks
        -------
        The `get_storage` method obtains a `StorageItem` on a `Folder` object
        using the identifier specified by `storage_identifier` and has the
        identifier type specified by `storage_identifier_type`. The
        `StorageItem` is a hidden item in the `Folder`, which roams with the
        account and is available online and offline.

        If you specify the `entry_id` for the `StorageItem` by using the
        `OlStorageIdentifierType.ENTRY_ID` value for `storage_identifier_type`,
        then the `get_storage` method will return the `StorageItem` with the
        specified `entry_id`. If no `StorageItem` can be found using that
        `entry_id` or if the `StorageItem` does not exist, then the
        `get_storage` method will raise an error.

        If you specify the message class for the `StorageItem` by using the
        `OlStorageIdentifierType.MESSAGE_CLASS` value for
        `storage_identifier_type`, then the `get_storage` method will return
        the `StorageItem` with the specified message class. If there are
        multiple items with the same message class, then the `get_storage`
        method returns the item with the most recent
        `PidTagLastModificationTime`. If no `StorageItem` exists with the
        specified message class, then the `get_storage` method creates a new
        `StorageItem` with the message class specified by `storage_identifier`.

        If you specify the `subject` of the `StorageItem`, then the
        `get_storage` method will return the `StorageItem` with the `subject`
        specified in the `get_storage` call. If there are multiple items with
        the same `subject`, then the `get_storage` method will return the item
        with the most recent `PidTagLastModificationTime`. If no `StorageItem`
        exists with the specified `subject`, then the `get_storage` method will
        create a new `StorageItem` with the `subject` specified by
        `storage_identifier`.

        `get_storage` returns an error if the store type of the folder is not
        supported.

        The `size` of a `StorageItem` that is newly created is zero (`0`) until
        you make an explicit call on the `save` method of the item.
        '''
        return self._folder.GetStorage(storage_identifier,
                                       storage_identifier_type)

    def get_table(
            self,
            filter_ = None,
            table_contents: int | _enums.OlTableContents = 0
    ) -> CDispatch:
        '''
        Obtains a `Table` object that contains items filtered by `filter_`.

        Parameters
        ----------
        filter_
            A filter in Microsoft Jet or DAV Searching and Locating (DASL)
            syntax that specifies the criteria for items in the parent
            `Folder`.
        table_contents : int | OlTableContents, default: USER_ITEMS
            Specifies the type of items in the folder that `get_table` returns.

        Returns
        -------
        CDispatch
            A `Table` that contains items in the parent `Folder` that meet the
            criteria in `filter_`. By default, `table_contents` is `USER_ITEMS`
            and the returned `Table` contains only the filtered items that are
            not hidden.

        Remarks
        -------
        If `filter_` is a blank string or the `filter_` parameter is omitted,
        `get_table` returns a `Table` with rows representing all the items in
        the `Folder`. If `filter_` is a blank string or the `filter_` parameter
        is omitted and `table_contents` is `HIDDEN_ITEMS`, `get_table` returns
        a `Table` with rows representing all the hidden items in the `Folder`.

        `get_table` returns a `Table` with the default column set for the
        folder type of the parent `Folder`. To modify the default column set,
        use the `add`, `remove`, or `remove_all` methods of the `Columns`
        collection object. When `table_contents` is `HIDDEN_ITEMS`, the default
        column set is always the default column set for a mail folder even
        though the parent `Folder` might be, for example, a Contacts folder.

        You can use `restrict` to apply subsequent filters to a Table that is
        based on the `Folder` object.
        '''
        return self._folder.GetTable(filter_, table_contents)

    def items(self) -> list[CDispatch]:
        '''
        Returns a list of `Item` objects from the specified folder. Read-only.

        Returns
        -------
        list[CDispatch]

        Remarks
        -------
        The `items` list is not guaranteed to be in any particular order.
        '''
        return list(self.data)

    def move_to(self, destination_folder: Folder) -> None:
        '''
        Moves a folder to the specified destination folder.

        Parameters
        ----------
        destination_folder : Folder
            The destination `Folder` for the `Folder` that is being moved.

        Remarks
        -------
        Setting the `REG_MULTI_SZ` value, `DisableCrossAccountCopy`, in
        `HKCU\\Software\\Microsoft\\Office\\14.0\\Outlook` in the Windows
        registry has the side effect of disabling this method.
        '''
        _dest_folder = destination_folder._folder
        self._folder.MoveTo(_dest_folder)
        return

    def set_custom_icon(self, picture: CDispatch) -> None:
        '''
        Sets a custom icon that is specified by `picture` for the folder.

        Parameters
        ----------
        picture : CDispatch
            Specifies the custom icon for the folder.

        Remarks
        -------
        The `StdPicture` object specified by `picture` must have its `type_`
        property equal to `PICTYPE_ICON` or `PICTYPE_BITMAP`. The icon or
        bitmap resource can have a maximum size of 32x32. Icons that are 16x16
        or 24x24 are also supported, and Microsoft Outlook can scale up a 16x16
        icon if Outlook is running in high Dots Per Inch (DPI) mode. Icons of
        other sizes cause `set_custom_icon` to return an error.

        You can set a custom icon for a search folder and for all folders that
        do not represent a default or a special folder. 

        You can only call `get_custom_icon` from code that runs in-process as
        Outlook. A `StdPicture` object cannot be marshaled across process
        boundaries. If you attempt to call `get_custom_icon` from
        out-of-process code, an exception occurs.

        The custom folder icon that this method provides does not persist
        beyond the running Outlook session. Add-ins therefore must set the
        custom folder icon every time that Outlook boots.

        The custom folder icon does not appear in other Exchange clients such
        as Outlook Web Access, nor does it appear in Outlook running on a
        Windows Mobile device.
        '''
        return self._folder.SetCustomIcon(picture)
    
    # Collection Non-Implementation
    
    def __setitem__(self, index: int, value: CDispatch) -> None:
        raise NotImplementedError('Setting/Deleting not implemented.')
    
    def __delitem__(self, index: int) -> None:
        raise NotImplementedError('Setting/Deleting not implemented.')