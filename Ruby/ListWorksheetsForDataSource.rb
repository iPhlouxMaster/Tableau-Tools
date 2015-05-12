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

$CSVFileName = 'TT-DataSourceWorksheets.csv'
$CSVHeader   = ["Rec #", "Workbook", "Data Source (UI)", "Worksheet" ]

def init
  puts  "\n\n\tLooking for Worksheets related to Data Sources.\n\n"
  print "\n\tIn directory? (.): "
  input = STDIN.gets.chomp
  dir   = if input == '' then '' else input + "/" end
  # -
  print "\n\tTWB(s) Named? (*): "
  input = STDIN.gets.chomp
  twb   = if input == '' then '*.twb' else input end
  # -
  print "\n\tData Sources? (*): "
  input = STDIN.gets.chomp
  $ds   = if input == '' then '.*' else input end
  # -
  puts "\n\n"
  # puts "\n\tLooking for Data Source(s) matching #{$ds}\/ in '#{path}' Workbooks\n\n\n"
  # -
  $csv = CSV.open($CSVFileName, "w") # do |csv|
  $csv << $CSVHeader
  # -
  dir + twb
end

dataSourceSheets = {}

$recNum = 0
def processTWB twbWithDir
  return unless twbWithDir =~ /.twb$/
  puts "\t#{twbWithDir}\n\t=============================="
  twb = Twb::Workbook.new twbWithDir
  $datasources = {}
  twb.worksheets.each do |ws|
    ws.datasources.each do |ds|
      if ds.uiname =~ /#{$ds}/i then loadSourceSheet(ds.uiname, ws.name) end
    end
  end
  $datasources.each do |dsn, sheets|
    puts "\n\t -- #{dsn}\n\t  |"
    sheets.each do |sheet|
      puts "\t  |-- #{sheet}"
      $csv << [ $recNum+=0, twb.name, dsn, sheet ]
    end
    puts "\n"
  end
end

def loadSourceSheet ds, sheet
  if $datasources[ds].nil? then $datasources[ds] = [] end
  $datasources[ds].push sheet
end


path = init
Dir.glob(path) {|twb| processTWB twb }

$csv.close unless $csv.nil?

puts "\n\n\tThat's all, Folks.\n\n"