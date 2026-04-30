"""
Cleanup script to automatically remove all test folders from Deleted Items (no confirmation)
"""
import sys
import os

# Set UTF-8 encoding for stdout
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

# Add parent directory to path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))

from backend.outlook_session.session_manager import OutlookSessionManager
from backend.outlook_session.folder_operations import FolderOperations

def cleanup_test_folders():
    """Remove all TestSource and TestDest folders from Deleted Items"""
    try:
        with OutlookSessionManager() as session:
            folder_ops = FolderOperations(session)
            
            # Get Deleted Items folder
            deleted_items = folder_ops.get_folder("Deleted Items")
            if not deleted_items:
                print("Error: Could not find Deleted Items folder")
                return 1
            
            # Collect all test folders
            test_folders = []
            for folder in deleted_items.Folders:
                folder_name = folder.Name
                if folder_name.startswith("TestSource_") or folder_name.startswith("TestDest_") or folder_name == "TestFolder":
                    test_folders.append(folder_name)
            
            if not test_folders:
                print("No test folders found in Deleted Items")
                return 0
            
            print(f"Found {len(test_folders)} test folders to remove")
            print(f"Starting automatic deletion...\n")
            
            # Delete folders
            deleted_count = 0
            failed_count = 0
            
            for i, folder_name in enumerate(test_folders, 1):
                try:
                    # Get the folder directly from deleted_items
                    folder_to_delete = None
                    for folder in deleted_items.Folders:
                        if folder.Name == folder_name:
                            folder_to_delete = folder
                            break
                    
                    if folder_to_delete:
                        folder_to_delete.Delete()
                        print(f"[{i}/{len(test_folders)}] OK Removed: {folder_name}")
                        deleted_count += 1
                    else:
                        print(f"[{i}/{len(test_folders)}] SKIP Not found: {folder_name}")
                        failed_count += 1
                except Exception as e:
                    print(f"[{i}/{len(test_folders)}] FAIL Failed to remove {folder_name}: {str(e)}")
                    failed_count += 1
            
            print(f"\n{'='*60}")
            print(f"Summary:")
            print(f"  Successfully deleted: {deleted_count}")
            print(f"  Failed: {failed_count}")
            print(f"{'='*60}")
            
            return 0
            
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(cleanup_test_folders())