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

# field nodes are children of Data Source nodes
$fieldNodeXpaths  = ["column",
                     "column-instance",
                     ".//metadata-record"
                   ]

$CSVFileName = 'TT-FieldAttributes.csv'
$CSVHeader   = ["Rec #",            "Workbook",
                "Data Source (UI)", "Data Source Hash",
                "Node Type",        "Field Name",
                "Attribute Name",   "Attribute Value"
               ]

$csv = CSV.open($CSVFileName, "w") # do |csv|
$csv << $CSVHeader

def processTWB twbWithDir
  puts "  -- #{twbWithDir}"
  twb = Twb::Workbook.new(twbWithDir)
  twb.datasources.each { |ds| emitAttributes(twb,ds) }
end

$recNum = 0
def emitAttributes(twb, ds)
  dsNode = ds.node
  $fieldNodeXpaths.each do |xpath|
    fieldNodes = dsNode.xpath(xpath)
    fieldNodes.each do |fieldNode|
      fieldName  = fieldNode.attribute('name')
      fieldName  = fieldName.text.gsub(/^\[/,'').gsub(/\]$/,'') unless fieldName.nil?
      attributes = fieldNode.attributes
      attributes.each do |attribute|
        $csv << [ $recNum+=1,     twb.name,
                  ds.uiname,      ds.connHash,
                  fieldNode.name, fieldName,
                  attribute[0],   attribute[1]
                ]
      end
    end
  end
end

path = if ARGV.empty? then '**/*.twb' else ARGV[0] end
Dir.glob(path) { |twb| processTWB twb }

$csv.close unless $csv.nil?
