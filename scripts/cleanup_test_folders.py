"""
Cleanup script to remove all test folders from Deleted Items
"""
import sys
import os

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
            
            print(f"Found {len(test_folders)} test folders to remove:")
            for name in test_folders:
                print(f"  - {name}")
            
            # Ask for confirmation
            response = input(f"\nDo you want to delete all {len(test_folders)} test folders? (yes/no): ")
            if response.lower() not in ['yes', 'y']:
                print("Cancelled.")
                return 0
            
            # Delete folders
            deleted_count = 0
            failed_count = 0
            
            for folder_name in test_folders:
                try:
                    folder_path = f"Deleted Items/{folder_name}"
                    result = folder_ops.remove_folder(folder_path)
                    print(f"✓ Removed: {folder_name}")
                    deleted_count += 1
                except Exception as e:
                    print(f"✗ Failed to remove {folder_name}: {str(e)}")
                    failed_count += 1
            
            print(f"\n{'='*60}")
            print(f"Summary:")
            print(f"  Successfully deleted: {deleted_count}")
            print(f"  Failed: {failed_count}")
            print(f"{'='*60}")
            
            return 0
            
    except Exception as e:
        print(f"Error: {str(e)}", file=sys.stderr)
        return 1

if __name__ == "__main__":
    sys.exit(cleanup_test_folders())