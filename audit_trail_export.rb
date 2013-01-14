require_relative 'db_reader'
require_relative 'excel_spreadsheet'

require 'pp'
require 'yaml'

def map_to_event_hash(data)
  events = []
  data.each do |row|

    event = {
        code: row[0],
        first_name: row[1],
        last_name: row[2],
        open_lock_type: row[3],
        open_lock_time: row[4],
        status: row[5],
        email: row[6],
        lock_name: row[7],
        lock_location: row[8],
        code_type: row[9]
    }

    events.push(event)
  end
  events
end

configuration = YAML.load_file('audit_trail_config.yml')

db_reader = DbReader.new(configuration['db_file'])
data = db_reader.fetch('SELECT SerialID, FirstName, LastName, OpenLockType, OpenLockTime, Status, Department, LockName, LockLocation, CodeType FROM RecordInfor')

events = map_to_event_hash(data)

spreadsheet = ExcelSpreadsheet.new
worksheet = spreadsheet.add_worksheet('Export')
worksheet.add_row(['First Name', 'Last Name', 'Email', 'Code', 'Lock Type', 'Lock', 'Lock Location', 'Time', 'Status', 'Code Type'])

events.each do |event|
  worksheet.add_row(
   [
      event[:first_name],
      event[:last_name],
      event[:email],
      event[:code],
      event[:open_lock_type],
      event[:lock_name],
      event[:lock_location],
      event[:open_lock_time],
      event[:status],
      event[:code_type]
   ]
  )
end

spreadsheet.save_and_close(configuration['export_directory'], configuration['filename_prefix'])