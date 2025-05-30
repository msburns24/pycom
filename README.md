# 📦 PyCOM

Provides a cleaner interface to Windows COM objects. Effectively a wrapper
over the `pywin32` module.

The goal of this project is to provide a means of communicating with the COM
Object API, but with excellent documentation throughout.

-------------------------------------------------------------------------------

## 📚 Table of Contents

- [📦 PyCOM](#-pycom)
  - [📚 Table of Contents](#-table-of-contents)
  - [✨ Features](#-features)
  - [💾 Installation](#-installation)
  - [✅ Todo](#-todo)
    - [General](#general)
    - [Classes](#classes)
    - [Enums](#enums)
    - [Exceptions](#exceptions)

-------------------------------------------------------------------------------

## ✨ Features

- ✅ Clean interface to Win32's COM object model
- 📝 Documentation consistent with the `Microsoft.Office.Interop.Outlook` API

-------------------------------------------------------------------------------

## 💾 Installation

```PowerShell
# Clone the repo
git clone https://github.com/msburns24/pycom.git
cd project-name

# Create a virtual environment (optional but recommended)
python -m venv venv
source venv/bin/activate   # On Windows use: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

-------------------------------------------------------------------------------

## ✅ Todo

### General

- [ ] Remap attributes to a pass-through to the `CDispatch` object where
  applicable.

### Classes

- [x] Account
- [ ] AddressEntry
- [ ] AddressList
- [x] Application
- [ ] AppointmentItem
- [ ] Automation
- [ ] Category
- [ ] COMAddIn
- [ ] ContactCard
- [ ] ContactItem
- [ ] DistListItem
- [ ] Explorer
- [ ] FileObject
- [ ] IAssistance
- [ ] Inbox
- [ ] Inspector
- [ ] OutlookItem
- [ ] OutlookItemType
- [ ] JournalItem
- [ ] LanguageSettings
- [ ] MailItem
- [ ] MobileItem
- [ ] NameSpace
- [ ] NoteItem
- [ ] PickerDialog
- [ ] PostItem
- [ ] Recipient
- [ ] Reference
- [ ] Reminder
- [ ] Search
- [ ] SharingItem
- [ ] Store
- [ ] TaskItem
- [ ] TimeZone


### Enums

- [ ] OlAddressEntryUserType
- [ ] OlSharingProvider

### Exceptions

- [ ] E_INVALIDARG
