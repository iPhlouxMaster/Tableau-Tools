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

$templateTwb  = 'Template.twb'
$templateDash = 'Template'
$twbAppend    = '_styled_'

puts "\n\n"
puts " Setting Workbook dashboard formatting, using the formatting"
puts " from the #{$templateDash} dashboard"
puts "   in the #{$templateTwb} workbook\n\n"

$csv = CSV.open("TT-FormattedWorkbooks.csv", "w")
$csv << ["Workbook","Dashboard"]

def loadTemplate
  return 'Template.twb not found' unless File.file?('Template.twb')
  twb  = Twb::Workbook.new('Template.twb')
  dash = twb.dashboards['Template']
  return 'Template dashboard not found' if dash.nil?
  style = dash.node.at_xpath('./style')
  return '  ERROR - no style available from Template dashboard.' if style.nil?
  puts "   Dashboard styling:"
  styleRules = style.xpath('./style-rule')
  if styleRules.empty?
    puts "\n\t  Template dashboard formatting is default style."
  else
    styleRules.each do |rule|
      puts "\n\t Element: #{rule['element']}"
      formats = rule.xpath('./format')
      formats.each do |f|
        puts sprintf("\t -- %-16s : %s \n", f['attr'], f['value'])
      end
    end
  end
  puts "\n"
  return style
end

def processTwbs
  path = if ARGV.empty? then '*.twb' else ARGV[0] end
  puts " Looking for TWBs using: #{ARGV[0]} \n\n"
  Dir.glob(path) do |fname|
    setTwbStyle(fname) unless fname.eql?($templateTwb) || !fname.end_with?('.twb')
  end
end

def setTwbStyle fname
  return if fname.eql?($templateTwb) || fname.include?($twbAppend + '.twb')
  twb = Twb::Workbook.new(fname)
  dashes = twb.dashboards.values
  puts sprintf("\t%3d in: '%s' ", dashes.length, fname)
  return if dashes.empty?
  dashes.each do |dash|
    node  = dash.node
    style = node.at_xpath('./style')
    tmpStyle = $templateStyle.clone
    style.replace(tmpStyle)
    $csv << [fname, dash.name]
  end
  twb.writeAppend($twbAppend)
end

$templateStyle = loadTemplate
if $templateStyle.class == 'String'.class
  puts "\t #{$templateStyle}\n\n"
else
  processTwbs
end

$csv.close unless $csv.nil?