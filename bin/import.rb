require_relative '../lib/access2000_database'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'
require 'csv'
require 'pathname'

def find_most_recent_backup_file(backup_file_table_name, import_directory)
  filenames = Dir[File.join(import_directory.gsub('\\', '/'), "*#{backup_file_table_name}.csv")]

  timestamps = []
  filenames.each do |filename|
    match = filename.match('.*/(.*)_backup.*')
    if match.length > 1
      timestamp = match[1]
      timestamps.push(timestamp)
    end
  end

  latest_filename = nil
  if timestamps.length > 0
    timestamps.sort!.reverse!

    filenames.each do |filename|
      match = filename.match('.*' + timestamps[0] + '.*')
      if !match.nil?
        latest_filename = match[0]
      end
    end
  end

  latest_filename
end

def restore_access_events(import_directory, db)
  backup_filename = find_most_recent_backup_file('access_events', import_directory)
  if !backup_filename.nil?
    puts "Importing access events from #{Pathname.new(backup_filename).basename}"
    db.delete_all_from_table('RecordInfor')

    counter = 0
    print "." # Print at least 1 dot even if there's less than 10 records to import
    STDOUT.flush
    CSV.foreach(backup_filename) do |row|
      if row[0] != 'SerialID' # Skip heading row
        counter += 1
        if counter % 10 == 0
          print "."
          STDOUT.flush
        end

        event = {}
        event[:serial_id] = row[0]
        event[:first_name] = row[1]
        event[:last_name] = row[2]
        event[:open_lock_type] = row[3].nil? ? '' : row[3]
        event[:open_lock_time] = row[4].nil? ? '' : row[4]
        event[:status] = row[5]
        event[:department] = row[6]
        event[:lock_name] = row[7].nil? ? '' : row[7]
        event[:lock_location] = row[8].nil? ? '' : row[8]
        event[:code_type] = row[9].nil? ? '' : row[9]

        insert_sql = 'INSERT INTO RecordInfor ' +
            "(SerialID, #{!event[:first_name].nil? ? 'FirstName, ' : ''}#{!event[:last_name].nil? ? 'LastName, ' : ''}OpenLockType, OpenLockTime, #{!event[:status].nil? ? 'Status, ' : ''}#{!event[:department].nil? ? 'Department, ' : ''}LockName, LockLocation, CodeType) " +
            "VALUES ('#{event[:serial_id]}', " +
            "#{!event[:first_name].nil? ? "'#{event[:first_name]}', " : ''}" +
            "#{!event[:last_name].nil? ? "'#{event[:last_name]}', " : ''}" +
            "'#{event[:open_lock_type]}', " +
            "'#{event[:open_lock_time]}', " +
            "#{!event[:status].nil? ? "'#{event[:status]}', " : ''}" +
            "#{!event[:department].nil? ? "'#{event[:department]}', " : ''}'" +
            "#{event[:lock_name]}', " +
            "'#{event[:lock_location]}', " +
            "'#{event[:code_type]}')"

        db.execute(insert_sql)
      end
    end

    puts
    puts "Imported #{counter} #{counter == 1 ? 'event' : 'events'}"
    puts
  else
    puts "An import file was not found for access events. Skipping."
  end
end

def restore_people(import_directory, db)
  backup_filename = find_most_recent_backup_file('people', import_directory)

  if !backup_filename.nil?
    puts "Importing people from #{Pathname.new(backup_filename).basename}"
    db.delete_all_from_table('KeyManagement')

    counter = 0
    print "." # Print at least 1 dot even if there's less than 10 records to import
    STDOUT.flush
    CSV.foreach(backup_filename) do |row|
      if row[0] != 'KeyID' # Skip heading row
        counter += 1
        if counter % 10 == 0
          print "."
          STDOUT.flush
        end

        person = {}
        person[:key_id] = row[0]
        person[:first_name] = row[1].nil? ? '' : row[1]
        person[:last_name] = row[2].nil? ? '' : row[2]
        person[:status] = row[3].nil? ? '' : row[3]
        person[:key_code] = row[4].nil? ? '' : row[4]
        person[:department] = row[5].nil? ? '' : row[5]
        person[:address] = row[6].nil? ? '' : row[6]
        person[:title] = row[7].nil? ? '' : row[7]
        person[:contact_infor] = row[8].nil? ? '' : row[8]
        person[:user_type] = row[9].nil? ? '' : row[9]
        person[:exp_date] = row[10].nil? ? '' : row[10]

        insert_sql = 'INSERT INTO KeyManagement ' +
            "(KeyID, FirstName, LastName, Status, KeyCode, Department, Address, Title, ContactInfor, UserType, ExpDate) " +
            "VALUES ('#{person[:key_id]}', '#{person[:first_name]}', '#{person[:last_name]}', '#{person[:status]}', " +
            "'#{person[:key_code]}', '#{person[:department]}', '#{person[:address]}', '#{person[:title]}', " +
            "'#{person[:contact_infor]}', '#{person[:user_type]}', #{person[:exp_date].to_s})"

        db.execute(insert_sql)
      end
    end

    puts
    puts "Imported #{counter} #{counter == 1 ? 'person' : 'people'}"
    puts
  else
    puts "An import file was not found for people. Skipping."
  end
end

def restore_locks(import_directory, db)
  backup_filename = find_most_recent_backup_file('locks', import_directory)

  if !backup_filename.nil?
    puts "Importing locks from #{Pathname.new(backup_filename).basename}"
    db.delete_all_from_table('LockManagement')

    counter = 0
    print "." # Print at least 1 dot even if there's less than 10 records to import
    STDOUT.flush
    CSV.foreach(backup_filename) do |row|
      if row[0] != 'LockID' # Skip heading row
        counter += 1
        if counter % 10 == 0
          print "."
          STDOUT.flush
        end

        lock = {}
        lock[:lock_id] = row[0]
        lock[:serial_id] = row[1].nil? ? '' : row[1]
        lock[:lock_name] = row[2].nil? ? '' : row[2]
        lock[:lock_location] = row[3].nil? ? '' : row[3]
        lock[:lock_status_id] = row[4]
        lock[:lock_type] = row[5].nil? ? '' : row[5]

        insert_sql = 'INSERT INTO LockManagement ' +
            "(LockID, SerialID, LockName, LockLocation, LockStatusID, LockType) " +
            "VALUES ('#{lock[:lock_id]}', '#{lock[:serial_id]}', '#{lock[:lock_name]}', '#{lock[:lock_location]}', " +
            "#{lock[:lock_status_id]}, '#{lock[:lock_type]}')"

        db.execute(insert_sql)
      end
    end

    puts
    puts "Imported #{counter} #{counter == 1 ? 'lock' : 'locks'}"
    puts
  else
    puts "An import file was not found for locks. Skipping."
  end
end

def restore_lock_assignments(import_directory, db)
  backup_filename = find_most_recent_backup_file('lock_assignments', import_directory)

  if !backup_filename.nil?
    puts "Importing the people assignments to each lock from #{Pathname.new(backup_filename).basename}"

    db.delete_all_from_table('LockKeyManagement')

    counter = 0
    print "." # Print at least 1 dot even if there's less than 10 records to import
    STDOUT.flush
    CSV.foreach(backup_filename) do |row|
      if row[0] != 'LockID' # Skip heading row
        counter += 1
        if counter % 10 == 0
          print "."
          STDOUT.flush
        end

        lock_key = {}
        lock_key[:lock_id] = row[0]
        lock_key[:key_id] = row[1]
        lock_key[:serial_id] = row[2].nil? ? '' : row[2]
        lock_key[:first_name] = row[3].nil? ? '' : row[3]
        lock_key[:last_name] = row[4].nil? ? '' : row[4]
        lock_key[:time_shift_id] = row[5]
        lock_key[:act_exp_date_id] = row[6]
        lock_key[:exp_date_id] = row[7]
        lock_key[:status] = row[8].nil? ? '' : row[8]
        lock_key[:department] = row[9].nil? ? '' : row[9]
        lock_key[:title] = row[10].nil? ? '' : row[10]
        lock_key[:address] = row[11].nil? ? '' : row[11]
        lock_key[:contact_infor] = row[12].nil? ? '' : row[12]
        lock_key[:last_setup_time] = row[13].nil? ? '' : row[13]
        lock_key[:user_type] = row[14].nil? ? '' : row[14]
        lock_key[:key_code] = row[15].nil? ? '' : row[15]
        lock_key[:exp_date] = row[16].nil? ? '' : row[16]

        insert_sql = 'INSERT INTO LockKeyManagement ' +
            "(LockID, KeyID, #{lock_key[:serial_id] == '' ? '' : "SerialID, "}FirstName, LastName," +
            " TimeShiftID, ActExpDateID, ExpDateID, Status," +
            " Department, Title, Address, ContactInfor, LastSetupTime, UserType, KeyCode, ExpDate) " +
            "VALUES (#{lock_key[:lock_id]}, " +
            "#{lock_key[:key_id]}, " +
            "#{lock_key[:serial_id] == '' ? '' : "'#{lock_key[:serial_id]}', "}" +
            "'#{lock_key[:first_name]}', " +
            "'#{lock_key[:last_name]}', " +
            "#{lock_key[:time_shift_id]}, " +
            "#{lock_key[:act_exp_date_id]}, " +
            "#{lock_key[:exp_date_id]}, " +
            "'#{lock_key[:status]}', " +
            "'#{lock_key[:department]}', " +
            "'#{lock_key[:title]}', " +
            "'#{lock_key[:address]}', " +
            "'#{lock_key[:contact_infor]}', " +
            "'#{lock_key[:last_setup_time]}', " +
            "'#{lock_key[:user_type]}', " +
            "'#{lock_key[:key_code]}', " +
            "#{lock_key[:exp_date].to_s.capitalize})"

        db.execute(insert_sql)
      end
    end

    puts
    puts "Imported #{counter} #{counter == 1 ? 'lock assignment' : 'lock assignments'}"
    puts
  else
    puts "An import file was not found for lock assignments. Skipping."
  end
end

config_file = File.join(File.dirname(__FILE__), '../config.yml')
dump_error_and_exit_if_file_missing(config_file)
configuration = YAML.load_file(config_file)
dump_error_and_exit_if_import_configuration_missing(config_file, configuration)

db = Access2000Database.new(configuration['db_file'])

restore_access_events(configuration['import_directory'], db)
restore_people(configuration['import_directory'], db)
restore_locks(configuration['import_directory'], db)
restore_lock_assignments(configuration['import_directory'], db)