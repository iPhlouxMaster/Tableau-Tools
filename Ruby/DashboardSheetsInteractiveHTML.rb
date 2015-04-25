# DashboardSheetsInteractiveHTML.rb - this Ruby script Copyright 2013, 2015 Christopher Gerrard

require 'nokogiri'

require 'twb'

def processTWB twbWithDir
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
  doc.title="Dashboards & their Worksheets for #{twbWithDir}"
  doc.write("#{twbWithDir}.dashboards.html")
end

system "cls"
puts "\n\n\tIdentifying the Worksheets in these Workbooks' Dashboards:"

path = if ARGV.empty? then '*.twb' else ARGV[0] end

Dir.glob(path) {|twb| processTWB twb }
