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

puts "\n\n"
puts " Setting Workbook dashboard formatting, using the formatting"
puts " from the #{$templateDash} dashboard"
puts "   in the #{$templateTwb} workbook\n\n"

$csv = CSV.open("TT-FormattedWorkbooks.csv", "w")
$csv << ["Workbook","Dashboard"]

def loadTemplate
  return 'Template.twb not found' unless File.file?('Template.twb')
  twb = Twb::Workbook.new('Template.twb')
  dash = twb.dashboards['Template']
  return 'Template dashboard not found' if dash.nil?
  styles = dash.node.xpath('./style/style-rule')
  return 'No styles available from Template dashboard' if styles.empty?
  puts "   Dashboard styling:"
  styleRules = {}
  styles.each do |rule|
    styleRules[rule['element']] = rule
    puts "\n\t Element: #{rule['element']}"
    formats = rule.xpath('./format')
    formats.each do |f|
      puts sprintf("\t -- %-16s : %s \n", f['attr'], f['value'])
    end
  end
  puts "\n"
  return styleRules
end

def setTwbStyle fname
  return if fname.eql? $templateTwb
  twb = Twb::Workbook.new(fname)
  dashes = twb.dashboards.values
  puts sprintf("\t%3d in '%s' ", dashes.length, fname)
  return if dashes.empty?
  dashes.each { |dash| setDashStyle(dash.node) } 
end

def setDashStyle dashNode
  puts "\t        - #{dashNode['name']}"
  $styleRules.each do |name, rule|
    setStyle(dashNode, name, rule)
  end
end

def setStyle(dashNode, name, rule)
  styleNode = dashNode.at_xpath('style')
  currStyle = dashNode.at_xpath(name)
  
  puts "\n\n++++"
  puts "name: #{name}"
  puts "rule: #{rule}"
  puts "curr: #{currStyle}"
end

def processTwbs
  path = if ARGV.empty? then '*.twb' else ARGV[0] end
  puts " Looking for TWBs using: #{ARGV[0]} .. \n\n"
  Dir.glob(path) do |fname|
    setTwbStyle(fname) unless fname.eql?($templateTwb) || !fname.end_with?('.twb')
  end
end

$styleRules = loadTemplate
if $styleRules.class == 'String'.class
  puts "\t #{$styleRules}\n\n"
else
  processTwbs
end

