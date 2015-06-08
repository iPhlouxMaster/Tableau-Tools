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

$CSVFileName = "TT-TableauFIlesXMLStructure.csv"
$CSVHeader   = ["Rec #","Workbook","Version","Depth","XML Path","XML Node","XML Attribute"]

def init
    $csv = CSV.open($CSVFileName,'w')
    $csv << $CSVHeader
    $recNum = 0
    $twbName = 0
end


def processTWB twbWithDir
  twb      = Twb::Workbook.new(twbWithDir)
  $twbName = $twbName += 1  #twb.name
  $version = twb.version
  depth   = 0
  puts "  -- #{$twbName}"
  parseNode(depth, '', twb.workbooknode)
end

def parseNode(depth, path,  node)
  currentDepth = depth += 1
  currentPath  = path + '\\' + node.name
  $csv << [$recNum +=1, $twbName, $version, depth, currentPath, node.name, ''] unless node.name.eql? 'text'
  attributes = node.attributes
  attributes.each do |name,value|
    $csv << [$recNum +=1, $twbName, $version, depth, currentPath, node.name, name] unless node.name.eql? 'text'
  end
  children = node.children
  #childrenPath = path +  node.name + '\\'
  children.each do |child|
    parseNode(currentDepth, currentPath, child)
  end
end

init
system 'cls'

puts "Looking for #{ARGV[0]}"
path = if ARGV.empty? then '**/*.twb' else ARGV[0] end
Dir.glob(path) {|twb| processTWB twb }

$csv.close unless $csv.nil?
