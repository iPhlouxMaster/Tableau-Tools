# ExtractFieldComments.rb - this Ruby script Copyright 2013, 2015 Chris Gerrard

require 'nokogiri'
require 'twb'
require 'csv'

$recNum = 0

$CSVFileName = "TT-FieldComments.csv"

$CSVHeader   = ["Rec #", "Workbook",
                "Data Source", "Data Source Hash",
                "Field - UI Name", "Field - Caption", "Field - Db Name", "Comments"]
def init
    $csv = CSV.open($CSVFileName,'w')
    $csv.puts $CSVHeader
end

def processTWB twbWithDir
  twb     = Twb::Workbook.new(twbWithDir)
  puts "  -- #{twb.name}"
  dataSources = twb.datasources
  dataSources.each do |ds|
    puts "\t == #{ds.uiname}"
    fields = ds.localfields
    fields.each do |name, field|
      comments = field.getComments
      unless comments.nil? || comments.eql?('')
        $csv << [$recNum+=1, twb.name,
                 ds.uiname,  ds.connHash,
                 field.uiname, field.caption, field.name, field.comments
                ]
      end
    end
  end
end

init

path = if ARGV.empty? then '**/*.twb' else ARGV[0] end
Dir.glob("*.twb") {|twb| processTWB twb }

$csv.close unless $csv.nil?
