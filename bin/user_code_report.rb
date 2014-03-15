require_relative '../lib/access2000_database'
require_relative '../lib/excel_spreadsheet'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'
require 'axlsx'

config_file = File.join(File.dirname(__FILE__), '../config.yml')
dump_error_and_exit_if_file_missing(config_file)
configuration = YAML.load_file(config_file)
dump_error_and_exit_if_user_code_configuration_missing(config_file, configuration)

db = Access2000Database.new(configuration['db_file'])
people_data = db.fetch('SELECT KeyID, FirstName, LastName, Status, KeyCode, Department, Address, Title, ContactInfor FROM KeyManagement')
lock_data = db.fetch('SELECT LockID, SerialID, LockName, LockLocation, LockStatusID, LockType FROM LockManagement')
people_lock_data = db.fetch('SELECT LockID, KeyID FROM LockKeyManagement')

people = map_to_people_hash(people_data)
locks = map_to_locks_hash(lock_data)
lock_assignments = map_to_lock_assignments_short_hash(people_lock_data)

people.sort! { |a, b| a[:last_name] <=> b[:last_name] }

if configuration.has_key?('ignore_these_locks')
  remove_ignored_locks(configuration['ignore_these_locks'], locks)
end
locks.sort! { |a, b| a[:name] <=> b[:name] }

Axlsx::Package.new :author => 'ABC Technical Support Group' do |spreadsheet|
  spreadsheet.workbook.use_autowidth = true
  spreadsheet.workbook.styles do |style|
    bold_style = style.add_style :b => true
    numeric_style = style.add_style :alignment => { :horizontal => :right }

    spreadsheet.workbook.add_worksheet(:name => 'Codes') do |worksheet|
      title_row_data = ['First Name', 'Last Name', 'Email', 'Phone', 'Suite', 'Status', 'Code']
      locks.each { |lock| title_row_data.push(lock[:name]) }
      worksheet.add_row(title_row_data, style: [bold_style] * title_row_data.length)

      people.each do |person|
        data_row = [
            person[:first_name],
            person[:last_name],
            person[:email],
            person[:phone],
            person[:suite],
            person[:status],
            person[:code]
        ]
        locks.each do |lock|
          assignment = lock_assignments.detect { |lock_assignment| lock_assignment[:lock_id] == lock[:id] && lock_assignment[:person_id] == person[:id] }
          if !assignment.nil?
            data_row.push(lock[:name])
          else
            data_row.push('')
          end
        end

        worksheet.add_row(data_row)
      end
    end
  end
  filename = File.join(configuration['export_directory'], configuration['user_code_filename_prefix'] + Time.now.strftime('%Y-%m-%d_%H-%M') + '.xlsx')
    spreadsheet.serialize(filename)
end