from __future__ import annotations
from rich import print, inspect
from src import Application


# Constants
INBOX_FOLDER_NUMBER = 6


def main() -> int:
    app = Application()
    namespace = app.get_namespace('MAPI')
    inbox = namespace.get_default_folder(INBOX_FOLDER_NUMBER)
    inspect(inbox)
    return 0



if __name__ == '__main__':
    error_code = main()
    exit(error_code)
