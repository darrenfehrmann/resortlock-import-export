call C:\Ruby193\bin\setrbvars.bat

echo Generating audit trail report
ruby C:\Users\edge\Documents\GitHub\resortlock-import-export\bin\audit_trail_report.rb

echo Generating user report
ruby C:\Users\edge\Documents\GitHub\resortlock-import-export\bin\user_code_report.rb

echo Generating studio membership report
ruby C:\Users\edge\Documents\GitHub\resortlock-import-export\bin\membership_report.rb

echo Backing up ResortLock database
ruby C:\Users\edge\Documents\GitHub\resortlock-import-export\bin\backup.rb

