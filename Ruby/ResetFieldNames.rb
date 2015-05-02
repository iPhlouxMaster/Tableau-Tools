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

$resetTWBextension = '.dashdoc'

$resets = []

def processTWB twbName
  return if twbName  =~ /dashdoc/
  puts "\t -- #{twbName}"
  twb   = Twb::Workbook.new(twbName)
  dirty = false
  dataSources = twb.datasources
  dataSources.each do |ds|
    if  ds.uiname != 'Parameters'
      fields = ds.localfields
      fields.each do |name,field|
        if (!field.caption.nil? || field.name.equal?( field.caption)) && field.calculation.nil?
          field.remove_attribute('caption')
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
