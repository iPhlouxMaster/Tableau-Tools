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
require 'csv'

$recNum = 0

$CSVFileName = "TT-FieldComments.csv"

$CSVHeader   = ["Rec #", "Workbook",
                "Data Source", "Data Source Hash",
                "Field - UI Name", "Field - Caption", "Field - Db Name", "Comments"]
def init
    $csv = CSV.open($CSVFileName,'w')
    $csv.puts $CSVHeader
end

def processTWB twbWithDir
  twb     = Twb::Workbook.new(twbWithDir)
  puts "  -- #{twb.name}"
  dataSources = twb.datasources
  dataSources.each do |ds|
    puts "\t == #{ds.uiname}"
    fields = ds.localfields
    fields.each do |name, field|
      comments = field.getComments
      unless comments.nil? || comments.eql?('')
        $csv << [$recNum+=1, twb.name,
                 ds.uiname,  ds.connHash,
                 field.uiname, field.caption, field.name, field.comments
                ]
      end
    end
  end
end

init

path = if ARGV.empty? then '**/*.twb' else ARGV[0] end
Dir.glob("*.twb") {|twb| processTWB twb }

$csv.close unless $csv.nil?
