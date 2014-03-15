require_relative '../lib/access2000_database'
#require_relative '../lib/excel_spreadsheet'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'
require 'axlsx'

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

Axlsx::Package.new :author => 'ABC Technical Support Group' do |spreadsheet|
  spreadsheet.workbook.use_autowidth = true
  spreadsheet.workbook.styles do |style|
    events_split_by_location.each_key do |lock_location|
      events_split_by_location[lock_location].sort! do |a, b|
        a[:open_lock_time] <=> b[:open_lock_time]
      end
      events_split_by_location[lock_location].reverse!  # Sort by most recent event at the top

      bold_style = style.add_style :b => true
      numeric_style = style.add_style :alignment => { :horizontal => :right }
      date_style = style.add_style :format_code => 'd-mmm-yyyy h:mm:ss AM/PM'
      spreadsheet.workbook.add_worksheet(:name => lock_location) do |worksheet|
        title_row_data = ['First Name', 'Last Name', 'Code', 'Lock Type', 'Lock', 'Lock Location', 'Time', 'Status', 'Code Type']
        worksheet.add_row(title_row_data, style: [bold_style] * title_row_data.length)

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
              ], style: [nil, nil, numeric_style, nil, nil, nil, date_style, nil, nil]
          )
        end
      end
    end
  end
  filename = File.join(configuration['export_directory'], configuration['audit_trail_filename_prefix'] + Time.now.strftime('%Y-%m-%d_%H-%M') + '.xlsx')
  spreadsheet.serialize(filename)
end
