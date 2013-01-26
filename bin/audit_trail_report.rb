require_relative '../lib/access2000_database'
require_relative '../lib/excel_spreadsheet'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'

config_file = File.join(File.dirname(__FILE__), '../config.yml')
dump_error_and_exit_if_file_missing(config_file)
configuration = YAML.load_file(config_file)
dump_error_and_exit_if_audit_trail_configuration_missing(config_file, configuration)

db = Access2000Database.new(configuration['db_file'])
data = db.fetch('SELECT SerialID, FirstName, LastName, OpenLockType, OpenLockTime, Status, Department, LockName, LockLocation, CodeType FROM RecordInfor')

events = map_to_event_hash(data)

events_split_by_location = {}
events.each do |event|
  if !events_split_by_location.has_key?(event[:lock_location])
    events_split_by_location[event[:lock_location]] = []
  end

  events_split_by_location[event[:lock_location]].push(event)
end

spreadsheet = ExcelSpreadsheet.new

events_split_by_location.each_key do |lock_location|
  events_split_by_location[lock_location].sort! do |a, b|
    a[:open_lock_time] <=> b[:open_lock_time]
  end
  events_split_by_location[lock_location].reverse!  # Sort by most recent event at the top

  worksheet = spreadsheet.add_worksheet(lock_location)
  worksheet.add_row(['First Name', 'Last Name', 'Code', 'Lock Type', 'Lock', 'Lock Location', 'Time', 'Status', 'Code Type'], {bold: true})

  events_split_by_location[lock_location].each do |event|
    worksheet.add_row(
     [
        event[:first_name],
        event[:last_name],
        event[:code],
        event[:open_lock_type],
        event[:lock_name],
        event[:lock_location],
        event[:open_lock_time],
        event[:status],
        event[:code_type]
     ], {columns_to_format_as_text: ['C']}
    )
  end
  worksheet.autofit(10)
end

spreadsheet.save_and_close(configuration['export_directory'], configuration['audit_trail_filename_prefix'])