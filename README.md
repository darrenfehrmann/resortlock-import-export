ResortLock Import/Export
========================

# Installation
1. On a windows machine, install [RubyInstaller](http://rubyinstaller.org/), for Ruby version 1.9.3.
2. Once Ruby is installed on your machine, download the latest version of this repository as a .zip file (look for
the Download ZIP link on the Github page).
3. Make sure the folder C:\Users\edge\Documents\GitHub\ exists.
4. Extract the .zip file to your Documents\Github folder on the Windows machine.
5. Inside the extracted files, delete the config.yml file and rename config.yml.laptop to config.yml to replace it.
This file configures the script.
6. Open config.yml (or view it online [here](https://github.com/darrenfehrmann/resortlock-import-export/blob/master/config.yml.laptop))
and make sure the export_directory, import_directory, and backup_directory all exist on the filesystem.
7. Open a command line window, cd to the ResortLock Import/Export directory, and run these two commands:
````
gem install bundler
bundle install
````
This will install the libraries that the scripts use to create .zip files and the .xlsx reports.
8. Create a shortcut to the run_reports_and_backups.bat file on the desktop.

## Other notes
If you ever want to link to one of the scripts directly, add this shortcut target to a new shortcut on the desktop:

````
C:\Windows\System32\cmd.exe /E:ON /K C:\Ruby193\bin\setrbvars.bat & ruby C:\Users\edge\Documents\GitHub\resortlock-import-export\bin\membership_report.rb & exit
````
and the "Start in" field to C:\Users\edge\Documents\