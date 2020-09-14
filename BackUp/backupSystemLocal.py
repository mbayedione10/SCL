#NioulBoy 07/2020
"""
    Backup script do the following:
    - Backups 7 most recent files specified in the list;
    - Removes the old archives in backup directory which exceed 10 retention days
    Required parameters:
    :param  BACKUP_DIR: Path to the backup directory. Creates if doesn't exists
    :param  TO_BACKUP: files/dirs paths to backup.
"""
import datetime
import zipfile
import os

BACKUP_DIR = r"C:\Users\NioulBoy\Documents\Backup"
TO_BACKUP = r"C:\Users\NioulBoy\Documents"


def create_backup_zip():
    """
    Creates zip archive with files from list in backup directory.
    - Print warning if dir/file in list doesn't exists.
    Args:
        list_of_files: Lists all the files in our backup directory ending with *.bak
        latest_file: Sorts the list of files by creation time, and return the 7 most recent
        zip_file: path of zip archive

    """
    date = datetime.datetime.now()
    timestamp = date.strftime("%Y-%m-%d")
    backup_name = "{0}-{1}{2}".format("backup", timestamp, ".zip")

    zip_file = zipfile.ZipFile(os.path.join(TO_BACKUP, backup_name), 'w')
    list_of_files = [os.path.join(BACKUP_DIR, file) for file in os.listdir(BACKUP_DIR) if file.endswith(".xlsx")]
    latest_file = sorted(list_of_files, key=os.path.getmtime, reverse=True)[:7]


    print("--------------------------Uploading------------------------")
    for file in latest_file:
        if os.path.isfile(file):
            zip_file.write(file)
            print("File " + file + " added Successfully")
        else:
            print("Warning! Path {0} doesn't exists. Please specify only exsisting paths.".format(latest_file[file]))
    zip_file.close()
    print("------The folder: " + backup_name + " has been created --------")


def remove_old_backup(backup_dir):
    """
    Removes the old archives in backup directory which exceed retention days.
    Args:
        backup_dir: Path to the backup directory.
    """

    date = datetime.datetime.today()
    old_date = date + datetime.timedelta(days=-10)
    for f in os.listdir(backup_dir):
         if datetime.datetime.fromtimestamp(os.path.getmtime(os.path.join(backup_dir,f))) < old_date:
             os.remove(os.path.join(backup_dir, f))
             print("Removing old backup {0} after 10 retention days".format(f))
         else:
             print("Old backup {0} is not older then 10 days. Skipping removal...".format(f))



def main():
    create_backup_zip()
    remove_old_backup(backup_dir=BACKUP_DIR)
main()
