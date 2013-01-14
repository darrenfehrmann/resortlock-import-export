require 'win32ole'

class DbReader
  def initialize(db_file_path)
    @db_file_path = db_file_path

    @db_password = get_access_2000_database_password
  end

  ##
  # Decodes the password from password-protected .mdb files that were created with Microsoft Access 2000.
  # Note that the password algorithm is specific to Access 2000 and this will not work for Access 97 or
  # 2003+ versions.
  #
  # The algorithm was converted from the VB code listed at:
  # http://tutorialsto.com/database/access/crack-access-*.-mdb-all-current-versions-of-the-password.html
  def get_access_2000_database_password()
  	password = ''

  	File.open(@db_file_path) do |file|
  		buffer = file.read(0x100)

  		encrypt_flag = buffer[0x62].ord

  		xor_bytes = [0xa1, 0xec, 0x7a, 0x9c, 0xe1, 0x28, 0x34, 0x8a, 0x73, 0x7b, 0xd2, 0xdf, 0x50]

  		xor_i = 0
  		file_i = 66
  		begin
  			if xor_i % 2 == 0
  				current_char = 0x13 ^ encrypt_flag ^ buffer[file_i].ord ^ xor_bytes[xor_i]
  			else
  				current_char = buffer[file_i].ord ^ xor_bytes[xor_i]
  			end

  			if current_char == 0
  				break
  			end
  			password = password + current_char.chr
  			file_i = file_i + 2
  			xor_i = xor_i + 1
  		end while xor_i < 13
  	end
  	password
  end

  def connect
    @connection = WIN32OLE.new('ADODB.Connection')
    @connection.open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=#{@db_file_path}; Jet OLEDB:Database Password=#{@db_password}")
  end

  def fetch(sql)
    connect

 		recordset = WIN32OLE.new('ADODB.Recordset')
 		recordset.open(sql, @connection)

 		data = recordset.GetRows.transpose
    recordset.close

    data
 	end
end