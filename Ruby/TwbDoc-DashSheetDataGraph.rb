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

puts "\n\n\t Documenting Workbooks with Tableau Tools"
puts   "\n\t Adding Dashboard -> Worksheet -> Data Source graphs"
puts   "\n\t https://github.com/ChrisGerrard/Tableau-Tools"
puts   "\n"

path = if ARGV.empty? then '*.twb' else ARGV[0] end
puts   "\n\t Files matching: '#{path}'"
Dir.glob(path) do |twb|
  puts "\t -- #{twb}"
  twb        = Twb::Workbook.new(twb)
  dotBuilder = Twb::Util::TwbDashSheetDataDotBuilder.new(twb)
  dotFile    = dotBuilder.dotFileName
  renderer   = Twb::Util::DotFileRenderer.new
  imageFile  = renderer.render(dotFile,'png')
  dash       = Twb::DocDashboardImageVert.new
  dash.image=(imageFile)
  dash.title=('Dashboards, Worksheets, and Data Sources')
  twb.addDocDashboard(dash)
  twb.writeAppend('dot')
end
