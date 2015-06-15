#  Copyright (C) 2012, 2015  Chris Gerrard
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

#  This tool renders the Dot files associated with their root Workbooks
#  and identified in the CSV File "TT-TwbAndDot.csv"

require 'twb'
require 'csv'

$gvDotLocation = 'C:\\tech\\graphviz\\Graphviz2.38\\bin\\dot.exe'

$cnt = 0
CSV.foreach("TT-TwbAndDot.csv") do |twb,dot|
  if $cnt > 0
    puts "#{$cnt} #{twb} \t dot:#{dot} "
    system "\"#{$gvDotLocation}\" -Tpng -o\"#{twb}.png\" \"#{dot}\""
  end
  $cnt += 1
end