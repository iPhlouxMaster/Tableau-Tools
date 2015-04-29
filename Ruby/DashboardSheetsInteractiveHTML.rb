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

require 'nokogiri'

require 'twb'

def processTWB twbWithDir
  return if twbWithDir  =~ /dashdoc/
  puts "\t -- #{twbWithDir}"
  twb    = Twb::Workbook.new(twbWithDir)
  sheets = twb.worksheetNames
  dashes = twb.dashboardNames
  dashhash = {}
  twb.dashboards.each do |dsh|
    dashsheets = nil
    if dsh.worksheets
      dashsheets = []
      dsh.worksheets.each do |sheet|
        dashsheets.push(sheet.name) unless sheet.nil?
      end
    end
    dashhash[dsh.name] = dashsheets
  end
  doc = Twb::HTMLListCollapsible.new(dashhash)
  doc.title="#{twbWithDir}"
  doc.write("#{twbWithDir}.dashboards.html")
end

system "cls"
puts "\n\n\tIdentifying the Worksheets in these Workbooks' Dashboards:"

path = if ARGV.empty? then '*.twb' else ARGV[0] end

Dir.glob(path) {|twb| processTWB twb }
