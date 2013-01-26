require_relative '../lib/access2000_database'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'
require 'csv'

def csv_filename(entity_to_backup, current_time, backup_directory)
  File.join(backup_directory, current_time.strftime('%Y-%m-%d_%H-%M') + '_backup_' + entity_to_backup + '.csv')
end

config_file = File.join(File.dirname(__FILE__), '../config.yml')
dump_error_and_exit_if_file_missing(config_file)
configuration = YAML.load_file(config_file)
dump_error_and_exit_if_backup_configuration_missing(config_file, configuration)

db = Access2000Database.new(configuration['db_file'])
people_data = db.fetch('SELECT KeyID, FirstName, LastName, Status, KeyCode, Department, Address, Title, ContactInfor FROM KeyManagement')
lock_data = db.fetch('SELECT LockID, SerialID, LockName, LockLocation, LockStatusID, LockType FROM LockManagement')
people_lock_data = db.fetch('SELECT LockID, KeyID, SerialID, FirstName, LastName, TimeShiftID, ActExpDateID, ExpDateID, Status, Department, Title, Address, ContactInfor, LastSetupTime, UserType, KeyCode, ExpDate FROM LockKeyManagement')
event_data = db.fetch('SELECT SerialID, FirstName, LastName, OpenLockType, OpenLockTime, Status, Department, LockName, LockLocation, CodeType FROM RecordInfor')

people = map_to_people_hash(people_data)
locks = map_to_locks_hash(lock_data)
lock_assignments = map_to_lock_assignments_long_hash(people_lock_data)
events = map_to_event_hash(event_data)

current_time = Time.now

# Export people as CSV
CSV.open(csv_filename('people', current_time, configuration['backup_directory']), "wb") do |csv|
  heading_row = %w{KeyID FirstName LastName Status KeyCode Department Address Title ContactInfor}
  csv << heading_row
  
  people.each do |person|
    data_row = [
                person[:id],
                person[:first_name],
                person[:last_name],
                person[:status],
                person[:code],
                person[:email],
                person[:suite],
                person[:phone],
                ''
            ]
    csv << data_row
  end
end

# Export locks as CSV
CSV.open(csv_filename('locks', current_time, configuration['backup_directory']), "wb") do |csv|
  heading_row = %w{LockID SerialID LockName LockLocation LockStatusID LockType}
  csv << heading_row
  
  locks.each do |lock|
    data_row = [
                lock[:id],
                lock[:serial_id],
                lock[:name],
                lock[:location],
                lock[:status_id],
                lock[:type]
            ]
    csv << data_row
  end
end

# Export lock assignments as CSV
CSV.open(csv_filename('lock_assignments', current_time, configuration['backup_directory']), "wb") do |csv|
  heading_row = %w{LockID KeyID SerialID FirstName LastName TimeShiftID ActExpDateID ExpDateID Status Department Title Address ContactInfor LastSetupTime UserType KeyCode ExpDate}
  csv << heading_row
  
  lock_assignments.each do |lock_assignment|
    data_row = [
                lock_assignment[:lock_id],
                lock_assignment[:person_id],
                lock_assignment[:serial_id],
                lock_assignment[:first_name],
                lock_assignment[:last_name],
                lock_assignment[:time_shift_id],
                lock_assignment[:act_exp_date_id],
                lock_assignment[:exp_date_id],
                lock_assignment[:status],
                lock_assignment[:department],
                lock_assignment[:title],
                lock_assignment[:address],
                lock_assignment[:contact_infor],
                lock_assignment[:last_setup_time].nil? ? nil : lock_assignment[:last_setup_time].strftime("%m/%d/%Y %l:%M:%S %p").gsub('  ', ' '),
                lock_assignment[:user_type],
                lock_assignment[:key_code],
                lock_assignment[:exp_date]
            ]
    csv << data_row
  end
end

# Export access events as CSV
CSV.open(csv_filename('access_events', current_time, configuration['backup_directory']), "wb") do |csv|
  heading_row = %w{SerialID FirstName LastName OpenLockType OpenLockTime Status Department LockName LockLocation CodeType}
  csv << heading_row
  
  events.each do |event|
    data_row = [
                event[:code],
                event[:first_name],
                event[:last_name],
                event[:open_lock_type],
                event[:open_lock_time].strftime("%m/%d/%Y %l:%M:%S %p").gsub('  ', ' '),
                event[:status],
                event[:email],
                event[:lock_name],
                event[:lock_location],
                event[:code_type]
            ]
    csv << data_row
  end
end