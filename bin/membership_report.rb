require_relative '../lib/access2000_database'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'
require 'axlsx'

def find_people_assigned_to_lock(lock, lock_assignments, people)
  assignees = []
  lock_assignments.each do |lock_assignment|
    next if lock[:id] != lock_assignment[:lock_id]

    assignee = people.detect { |person| person[:id] == lock_assignment[:person_id] }

    if !assignee.nil?
      assignees.push(assignee)
    end
  end

  assignees.sort! { |a, b| a[:last_name] <=> b[:last_name] }
  assignees
end

config_file = File.join(File.dirname(__FILE__), '../config.yml')
dump_error_and_exit_if_file_missing(config_file)
configuration = YAML.load_file(config_file)
dump_error_and_exit_if_membership_configuration_missing(config_file, configuration)

db = Access2000Database.new(configuration['db_file'])
people_data = db.fetch('SELECT KeyID, FirstName, LastName, Status, KeyCode, Department, Address, Title, ContactInfor FROM KeyManagement')
lock_data = db.fetch('SELECT LockID, SerialID, LockName, LockLocation, LockStatusID, LockType FROM LockManagement')
people_lock_data = db.fetch('SELECT LockID, KeyID FROM LockKeyManagement')

people = map_to_people_hash(people_data)
locks = map_to_locks_hash(lock_data)
lock_assignments = map_to_lock_assignments_short_hash(people_lock_data)

if configuration.has_key?('ignore_these_locks')
  remove_ignored_locks(configuration['ignore_these_locks'], locks)
end

# Reverse sort the locks since the worksheets will be added in reverse order
locks.sort! { |a, b| b[:name] <=> a[:name] }

Axlsx::Package.new :author => 'ABC Technical Support Group' do |spreadsheet|
  spreadsheet.workbook.use_autowidth = true
  spreadsheet.workbook.styles do |style|
    bold_style = style.add_style :b => true
    numeric_style = style.add_style :alignment => { :horizontal => :right }
    title_row_style = style.add_style :sz => 18, :b => true

    locks.each do |lock|
      assignees = find_people_assigned_to_lock(lock, lock_assignments, people)

      spreadsheet.workbook.add_worksheet(:name => lock[:name]) do |worksheet|
        worksheet.add_row([lock[:location].to_s.upcase], style: [title_row_style])
        worksheet.merge_cells "A1:F1"
        worksheet.add_row(['First Name', 'Last Name', 'Email', 'Phone', 'Suite', 'Status'], style: [bold_style] * 6)

        assignees.each do |assignee|
          worksheet.add_row(
              [
                  assignee[:first_name],
                  assignee[:last_name],
                  assignee[:email],
                  assignee[:phone],
                  assignee[:suite],
                  assignee[:status]
              ], style: [nil, nil, nil, nil, numeric_style, nil] )
        end
      end
    end
  end
  filename = File.join(configuration['export_directory'], configuration['membership_filename_prefix'] + Time.now.strftime('%Y-%m-%d_%H-%M') + '.xlsx')
  spreadsheet.serialize(filename)
end