require_relative '../lib/access2000_database'
require_relative '../lib/entity_mappings'
require_relative '../lib/error_checking'

require 'pp'
require 'yaml'

config_file = File.join(File.dirname(__FILE__), '../config.yml')
dump_error_and_exit_if_file_missing(config_file)
configuration = YAML.load_file(config_file)
dump_error_and_exit_if_backup_configuration_missing(config_file, configuration)

db = Access2000Database.new(configuration['db_file'])

before_data = db.fetch('SELECT SerialID, FirstName, LastName, OpenLockType, OpenLockTime, Status, Department, LockName, LockLocation, CodeType FROM RecordInfor')

sql = "DELETE FROM RecordInfor WHERE OpenLockTime < #" + Time.now.strftime("%m") + "/01/" + (Time.now.year - 1).to_s + "#"
db.execute(sql)

after_data = db.fetch('SELECT SerialID, FirstName, LastName, OpenLockType, OpenLockTime, Status, Department, LockName, LockLocation, CodeType FROM RecordInfor')

number_records_deleted = before_data.length - after_data.length
puts number_records_deleted.to_s + " events occurring before " + Time.now.strftime("%b") + " 1, " + (Time.now.year - 1).to_s + " have been deleted"