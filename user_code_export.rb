require_relative 'db_reader'
require_relative 'excel_spreadsheet'

require 'pp'
require 'yaml'

def map_to_people_hash(data)
  people = []
  data.each do |row|

    person = {
        id: row[0],
        first_name: row[1].strip,
        last_name: row[2].strip,
        status: row[3].strip,
        code: row[4].strip,
        email: row[5].strip,
        suite: row[6].strip
    }
    if row[7] != nil && !row[7].empty?
      person[:phone] = row[7].strip
    elsif row[8] != nil && !row[8].empty?
      person[:phone] = row[8].strip
    else
      person[:phone] = ''
    end

    people.push(person)
  end

  people
end

def map_to_locks_hash(data)
  locks = []
  data.each do |row|

    lock = {
        id: row[0],
        serial_id: row[1].strip,
        name: row[2].strip,
        location: row[3].strip,
        status_id: row[4],
        type: row[5].strip
    }

    locks.push(lock)
  end

  locks
end

def map_to_lock_assignments_hash(data)
  lock_assignments = []
  data.each do |row|

    lock_assignment = {
        lock_id: row[0],
        person_id: row[1]
    }

    lock_assignments.push(lock_assignment)
  end

  lock_assignments
end

def remove_ignored_locks(lock_names, locks)
  lock_names.each do |name|
    locks.delete_if { |lock| lock[:name] == name }
  end
end

configuration = YAML.load_file('config.yml')

db_reader = DbReader.new(configuration['db_file'])
people_data = db_reader.fetch('SELECT KeyID, FirstName, LastName, Status, KeyCode, Department, Address, Title, ContactInfor FROM KeyManagement')
lock_data = db_reader.fetch('SELECT LockID, SerialID, LockName, LockLocation, LockStatusID, LockType FROM LockManagement')
people_lock_data = db_reader.fetch('SELECT LockID, KeyID FROM LockKeyManagement')

people = map_to_people_hash(people_data)
locks = map_to_locks_hash(lock_data)
lock_assignments = map_to_lock_assignments_hash(people_lock_data)

people.sort! { |a, b| a[:last_name] <=> b[:last_name] }

remove_ignored_locks(configuration['ignore_these_locks'], locks)
locks.sort! { |a, b| a[:name] <=> b[:name] }

spreadsheet = ExcelSpreadsheet.new

worksheet = spreadsheet.add_worksheet('Codes')
heading_row = ['First Name', 'Last Name', 'Email', 'Phone', 'Suite', 'Status', 'Code']
locks.each { |lock| heading_row.push(lock[:name]) }
worksheet.add_row(heading_row)

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

worksheet.autofit(7+locks.length)

spreadsheet.save_and_close(configuration['export_directory'], configuration['user_code_filename_prefix'])