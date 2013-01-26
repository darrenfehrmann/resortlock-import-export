def map_to_people_hash(data)
  people = []
  data.each do |row|

    person = {
        id: row[0],
        first_name: row[1].nil? ? '' : row[1].strip,
        last_name: row[2].nil? ? '' : row[2].strip,
        status: row[3].strip,
        code: row[4].strip,
        email: row[5].nil? ? '' : row[5].strip,
        suite: row[6].nil? ? '' : row[6].strip
    }
    if row[7] != nil && !row[7].empty?
      person[:phone] = row[7].strip
    elsif row[8] != nil && !row[8].empty?
      person[:phone] = row[8].strip
    else
      person[:phone] = ''
    end

    if row[9] != nil && !row[9].empty?
      person[:user_type] = row[9].strip
    else
      person[:user_type] = ''
    end

    if row[10] != nil
      person[:exp_date] = row[10]
    else
      person[:exp_date] = false
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

def map_to_lock_assignments_short_hash(data)
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

##
# The LockKeyManagement table holds a duplicate set of user data for each user. Normally we don't need this
# information, but for backups, we export the entire set of fields.
def map_to_lock_assignments_long_hash(data)
  lock_assignments = []
  data.each do |row|

    lock_assignment = {
        lock_id: row[0],
        person_id: row[1],
        serial_id: row[2],
        first_name: row[3],
        last_name: row[4],
        time_shift_id: row[5],
        act_exp_date_id: row[6],
        exp_date_id: row[7],
        status: row[8],
        department: row[9],
        title: row[10],
        address: row[11],
        contact_infor: row[12],
        last_setup_time: row[13],
        user_type: row[14],
        key_code: row[15],
        exp_date: row[16]
    }

    lock_assignments.push(lock_assignment)
  end

  lock_assignments
end

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

def remove_ignored_locks(lock_names, locks)
  lock_names.each do |name|
    locks.delete_if { |lock| lock[:name] == name }
  end
end
