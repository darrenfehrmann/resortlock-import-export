def dump_error_and_exit_if_file_missing(file)
  if !File.exists? file
    show_error_heading
    file_full_path = File.expand_path(file)
    puts
    puts "I can't find the configuration file"
    puts file_full_path
    puts
    puts "The configuration file is needed to run any of the reports."
    puts
    puts "To fix the problem, check to see if there is a file named"
    puts "config.yml.amenities_building. If so then rename the"
    puts "file config.yml.amenities_building to config.yml and run the report again."
    puts
    puts "If you cannot see any config.yml.* files then:"
    puts "1. Open up the Github project in your browser:"
    puts "   https://github.com/darrenfehrmann/resortlock-import-export"
    puts "2. You should see a list of files. Click on the"
    puts "   config.yml.amenities_building file to see the contents."
    puts "3. Create a new file called config.yml and copy-paste the contents"
    puts "   from the Github file into your new file."

    exit 1
  end
end

def show_error_heading
  puts
  puts "----------------------- Error -----------------------"
end

def show_error_if_missing_key(key, sample_line, configuration_file, configuration)
  if !configuration.has_key? key
    show_error_heading
    file_full_path = File.expand_path(configuration_file)
    puts
    puts "The configuration directive '#{key.to_s}' is missing from"
    puts file_full_path
    puts
    puts "To fix the problem, open the configuration file and add the line:"
    puts sample_line
    puts
    puts "The line above is just an example. The value can be whatever you like."
  end

  !configuration.has_key? key
end

def show_error_if_missing_db_config(configuration_file, configuration)
  if !configuration.has_key? 'db_file'
    show_error_heading
    file_full_path = File.expand_path(configuration_file)
    puts
    puts "The configuration directive 'db_file' is missing from"
    puts file_full_path
    puts
    puts "To fix the problem, open the configuration file and add the line:"
    puts "db_file: C:\\Program Files\\ResortLock3.7\\iKeypadLockDB.mdb"
    puts
    puts "This ought to work, but if your ResortLock installation is located"
    puts "somewhere else, then you'll have to change the path to point to the .mdb file."
  end

  !configuration.has_key? 'db_file'
end

def dump_error_and_exit_if_backup_configuration_missing(configuration_file, configuration)
  fail = false
  fail = true if show_error_if_missing_key('backup_directory', 'backup_directory: C:\Users\Darren\Documents\Excel\resortlock-import-export\backups', configuration_file, configuration)

  if fail
    exit 2
  end
end

def dump_error_and_exit_if_audit_trail_configuration_missing(configuration_file, configuration)
  fail = false
  fail = true if show_error_if_missing_db_config(configuration_file, configuration)
  fail = true if show_error_if_missing_key('export_directory', 'export_directory: C:\\Users\\edge\\Documents\\resortlock-import-export\\exports', configuration_file, configuration)
  fail = true if show_error_if_missing_key('audit_trail_filename_prefix', 'audit_trail_filename_prefix: audit_trail_', configuration_file, configuration)

  if fail
    exit 2
  end
end

def dump_error_and_exit_if_user_code_configuration_missing(configuration_file, configuration)
  fail = false
  fail = true if show_error_if_missing_db_config(configuration_file, configuration)
  fail = true if show_error_if_missing_key('export_directory', 'export_directory: C:\\Users\\edge\\Documents\\resortlock-import-export\\exports', configuration_file, configuration)
  fail = true if show_error_if_missing_key('user_code_filename_prefix', 'user_code_filename_prefix: user_code_', configuration_file, configuration)

  if fail
    exit 2
  end
end

def dump_error_and_exit_if_membership_configuration_missing(configuration_file, configuration)
  fail = false
  fail = true if show_error_if_missing_db_config(configuration_file, configuration)
  fail = true if show_error_if_missing_key('export_directory', 'export_directory: C:\\Users\\edge\\Documents\\resortlock-import-export\\exports', configuration_file, configuration)
  fail = true if show_error_if_missing_key('membership_filename_prefix', 'membership_filename_prefix: membership_', configuration_file, configuration)

  if fail
    exit 2
  end
end

def dump_error_and_exit_if_import_configuration_missing(configuration_file, configuration)
  fail = false
  fail = true if show_error_if_missing_db_config(configuration_file, configuration)
  fail = true if show_error_if_missing_key('import_directory', 'import_directory: C:\\Users\\edge\\Documents\\resortlock-import-export\\imports', configuration_file, configuration)

  if fail
    exit 2
  end
end