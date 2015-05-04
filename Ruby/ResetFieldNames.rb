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

# to overwrite the original TWB, set $resetTWBextension = ''
$resetTWBextension = '.reset'

$resets = []

# Accept a Workbook name and process it, removing field captions.
#
# The intent is to reverse Tableau's practice of altering the field names
# presented to Users in the Data window. When the names presented to the
# User are different from the database names Tableau uses a "caption" attribute
# to record the User name, i.e. it's the presence of the caption that signals
# Tableau to display the caption instead of the database field name.
#
# When Tableau initially opens a data connection and identifies field names that
# it determines need adjusting, it does at least these two things:
# - Replace underscores with blanks &ndash; 'this_is_a_field' -> 'this is a field'
# - Proper noun casing &ndash; 'this is a field' -> 'This Is A Field'
#
# While this is often desirable, there are many instances where it's contrary to
# the intentions of the Tableau user and the data managers.
# This script resets those field names that have been captioned to their original
# database names by removing the captions, in which case Tableau uses the database
# name for the field in the data window.
#
# Note: Calculated fields have captions but should not be de-captioned since they're
#       not actual database fields.
#
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

puts "\n\n\tReset #{$resets.length} fields' names."
puts   "\n\tThe reset fields are recorded in \"ResetFields.csv\"\n\n"
