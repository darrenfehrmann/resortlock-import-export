require_relative '../lib/access2000_database'
require_relative '../lib/excel_spreadsheet'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'

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

spreadsheet = ExcelSpreadsheet.new

locks.each do |lock|
  assignees = find_people_assigned_to_lock(lock, lock_assignments, people)

  worksheet = spreadsheet.add_worksheet(lock[:name])
  worksheet.add_row([lock[:location]], {bold: true, font_size: 18, merge_columns: 6, all_caps: true})
  worksheet.add_row(['First Name', 'Last Name', 'Email', 'Phone', 'Suite', 'Status'], {bold: true})

  assignees.each do |assignee|
    worksheet.add_row(
        [
            assignee[:first_name],
            assignee[:last_name],
            assignee[:email],
            assignee[:phone],
            assignee[:suite],
            assignee[:status]
        ])
  end
  worksheet.autofit(6)
end

spreadsheet.save_and_close(configuration['export_directory'], configuration['membership_filename_prefix'])