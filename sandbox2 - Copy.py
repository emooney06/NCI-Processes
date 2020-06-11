# There are multiple ways of navigating the folder tree and searching for folders. Globbing and 
# absolute path may create unexpected results if your folder names contain slashes.

# The folder structure is cached after first access to a folder hierarchy. This means that external
# changes to the folder structure will not show up until you clear the cache. Here's how to clear
# the cache of each of the currently supported folder hierarchies:
from exchangelib import Account, Folder

a = Account(...)
a.root.refresh()
a.public_folders_root.refresh()
a.archive_root.refresh()

some_folder = a.root / 'Some Folder'
some_folder.parent
some_folder.parent.parent.parent
some_folder.root  # Returns the root of the folder structure, at any level. Same as Account.root
some_folder.children  # A generator of child folders
some_folder.absolute  # Returns the absolute path, as a string
some_folder.walk()  # A generator returning all subfolders at arbitrary depth this level
# Globbing uses the normal UNIX globbing syntax, but case-insensitive
some_folder.glob('foo*')  # Return child folders matching the pattern
some_folder.glob('*/foo')  # Return subfolders named 'foo' in any child folder
some_folder.glob('**/foo')  # Return subfolders named 'foo' at any depth
some_folder / 'sub_folder' / 'even_deeper' / 'leaf'  # Works like pathlib.Path
# You can also drill down into the folder structure without using the cache. This works like
# the single slash syntax, but does not start by creating a cache the folder hierarchy. This is
# useful if your account contains a huge number of folders, and you already know where to go.
some_folder // 'sub_folder' // 'even_deeper' // 'leaf'
some_folder.parts  # returns some_folder and all its parents, as Folder instances
# tree() returns a string representation of the tree structure at the given level
print(a.root.tree())
'''
root
├── inbox
│   └── todos
└── archive
    ├── Last Job
    ├── exchangelib issues
    └── Mom
'''

# Folders have some useful counters:
a.inbox.total_count
a.inbox.child_folder_count
a.inbox.unread_count
# Update the counters
a.inbox.refresh()

# Folders can be created, updated and deleted:
f = Folder(parent=a.inbox, name='My New Folder')
f.save()

f.name = 'My New Subfolder'
f.save()
f.delete()

# Delete all items in a folder
f.empty()
# Also delete all subfolders in the folder
f.empty(delete_sub_folders=True)
# Recursively delete all items in a folder, and all subfolders and their content. This is
# like `empty(delete_sub_folders=True)` but attempts to protect distinguished folders from
# being deleted. Use with caution!
f.wipe()
