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

$resetTWBextension = '.reset'

$resets = []

# Accept a Workbook name and process it, removing field caption.
# The intent is to reverse Tableau's practice of altering the names
# presented to Users in the Data window.
# Tableau does two primary things
# - Replace underscores with blanks &ndash; 'this_is_a_field' -> 'this is a field'
# - Proper noun casing &ndash; 'this is a field' -> 'This Is A Field'
# While this is often desirable, there are many instances where it's contrary to
# the intentions of the Tableau user and the data managers.
def processTWB twbName
  return if twbName  =~ /$resetTWBextension/
  puts "\t -- #{twbName}"
  twb   = Twb::Workbook.new(twbName)
  dirty = false
  dataSources = twb.datasources
  dataSources.each do |ds|
    if  ds.uiname != 'Parameters'
      fields = ds.localfields
      fields.each do |name,field|
        if (   !field.caption.nil? || field.name.equal?( field.caption)) && field.calculation.nil?
          field.remove_attribute('caption')
          $resets << '"' + twbName + '","' + ds.name + '","' +  field.name + '","' +  field.caption + '"'
          dirty = true
        end
      end
    end
  end
  twb.writeAppend($resetTWBextension) if dirty
end

system "cls"
puts "\n\n\tResetting the Field names in these Workbooks:\n"

path = if ARGV.empty? then '*.twb' else ARGV[0] end
Dir.glob(path) {|twb| processTWB twb }

if !$resets.empty?
   csv = File.open("ResetFields.csv", 'w')
   csv.puts "Workbook.Data Connection,Field Name,Field Caption"
   $resets.each { |resetRec| csv.puts resetRec }
   csv.close unless csv.nil?
end

puts "\n\n\t\tReset #{$resets.length} fields' names."
