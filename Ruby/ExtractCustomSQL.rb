#  Copyright (C) 2014, 2015  Chris Gerrard
#
#  This program is free software: you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation, either version 3 of the License, or
#  (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with this program.  If not, see <http://www.gnu.org/licenses/>.

require 'twb'
require 'nokogiri'
require 'csv'

$replaceCSQLFile   = true

$csv = CSV.open("TT-CustomSQL.csv", "w")
$csv << ["Tableau File","Type","Data Source","CSQL File","CSQL # Chars","CSQL # Lines"]

# Workbooks
def twbCustomSQL(twbName)
  return unless twbName =~ /^.*[.]twb$/
  twb = Twb::Workbook.new(twbName)
  puts "    -- #{twb.name}"
  begin
    dataSources = twb.datasources
    dataSources.each do |ds|
      customSQLNode = ds.node.xpath(".//relation[@name='TableauSQL']")
      customSQLNode.each do |node|
        puts "       -> #{ds.name}"
        writeCSQL(twbName,'TWB',ds.name,node,twb.name," == ",ds.name)
      end
    end
  rescue
    puts "         ** "
    puts "         ** FAILED "
    puts "         ** "
  end
end

# TDS files
def tdsCustomSQL(tdsName)
  return unless tdsName =~ /^.*[.]tds$/
  puts "       -> #{tdsName}"
  doc = Nokogiri::XML(open(tdsName))
  textRelationNodes = doc.xpath("//relation[@type='text']")
  textRelationNodes.each do |node|
    writeCSQL(tdsName,'TDS',tdsName,node,File.basename(tdsName),' == ','')
  end
end

def writeCSQL(tabFile,type,dsName,node,lead,sep,trail)
  csqlFname = buildName(lead,sep,trail)
  File.delete(csqlFname) if $replaceCSQLFile && File.exist?(csqlFname)
  file = File.open(csqlFname, 'w')
  file << node.text
  file.close
  writeCSV(tabFile,dsName,type,csqlFname,node.text)
end

def writeCSV(tabFile,dsName,type,csqlFile,csql)
  lines = csql.split(/\n/)
  $csv << [tabFile,type,dsName,csqlFile,csql.length,lines.length]
end

def buildName(lead,sep,trail)
  fname = lead + sep + trail + ".csql"
  fname.gsub(/[^a-zA-Z0-9_\-=. ]/,'')
end

system 'cls'

path = if ARGV.empty? then '**/*.twb' else ARGV[0] end
puts "\n  Processing files matching '#{path}'\n "
Dir.glob(path)  { |twb| twbCustomSQL twb }

path = if ARGV.empty? then '**/*.tds' else ARGV[0] end
puts "\n  Processing files matching '#{path}'\n"
Dir.glob(path)  { |tds| tdsCustomSQL tds }
