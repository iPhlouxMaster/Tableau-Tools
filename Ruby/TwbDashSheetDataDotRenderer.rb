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

require 'twb'
require 'csv'

#====================================================================
$dotHeader = <<DOTHEADER
digraph g {
    graph [rankdir="LR" splines="line"];
    node  [shape="box"  width="2"];

    "Tableau Tools generated Workbook map" [color="white" border="0"];
DOTHEADER
#====================================================================

#====================================================================
$dotFooter = <<DOTFOOTER
  }

   subgraph cluster_0 {
    color=white;
    node [shape="box3d"  style="filled" ];
    "Workbook" -> "Dashboard" -> "Worksheet" -> "Data Source"
    "Workbook"    "Dashboard";
    "Worksheet";
    "Data Source";
   }

// -------------------------------------------------------------
}
DOTFOOTER
#====================================================================

def init
  $sheetCnt, $twbCnt, $dashCnt = 0, 0, 0
  $csv = CSV.open('TT-TwbAndDot.csv','w')
  $csv << ['Workbook','Dot File']
  $gvDotLocation = 'C:\\tech\\graphviz\\Graphviz2.38\\bin\\dot.exe'
end


def processTWB twbName
   twb = Twb::Workbook.new twbName
   puts "  -- #{twb.name}"
   $worksheets  = twb.worksheetNames
   $datasources = twb.datasourceUINames
   pairs  = processDashboards twb
   pairs += processWorksheets twb
   pairs += processOrphans    twb
   dotFile = initDot twb
   buildBody(dotFile,pairs)
   sameRank(dotFile,[twb.name])
   sameRank(dotFile,twb.dashboards.keys)
   sameRank(dotFile,twb.worksheetNames)
   sameRank(dotFile,twb.datasourceNames)
   buildHeader(dotFile)
   closeDot(dotFile)
   $csv <<  [twbName, dotFile.to_path]
   image = renderDot(twbName, dotFile.to_path)
   addImageToTwb(twb, image)
   $twbCnt += 1
end

def renderDot(twb,dot)
  imageType  = '-Tpng'
  imageFile  = twb + '.png'
  imageParam = '-o' + imageFile
  system $gvDotLocation, imageType, imageParam, dot
  return imageFile
end

def addImageToTwb(twb, image)
  dash = Twb::DocDashboardImageVert.new
  dash.image=(image)
  dash.title=('Dashboards, Worksheets, and Data Sources')
  twb.addDocDashboard(dash)
  twb.writeAppend('dot')
end

def inject(twb, dashboard, htmlfilename)
  vDash = Twb::DocDashboardWebVert.new
  vDash.title=('Doc Dashboard: ' + sanitize(dashboard))
  vDash.url=($localurl  + '/' + htmlfilename)
  twb.addDocDashboard(vDash)
  if $replacetwb
    twb.write
  else
    twb.writeAppend($dashdoclbl)
  end
end

def initDot twb
  dotFile = File.open("#{twb.name}.dot",'w')
  dotFile.puts $dotHeader
  return dotFile
end

def buildBody(dotFile,pairs)
  dotFile.puts "\n   subgraph cluster_1 {"
  dotFile.puts "       color= grey;"
  dotFile.puts ""
  pairs.each { |pair| dotFile.puts "      \"#{pair[0]}\" -> \"#{pair[1]}\" " }
  dotFile.puts ""
  dotFile.puts "   }"
end

def sameRank(dotFile, elements)
  dotFile.puts "\n  {rank=same "
  elements.each do |e|
    dotFile.puts "     \"#{e}\""
  end
  dotFile.puts "  }"
end


def buildHeader dotFile
  dotFile.puts ''
  dotFile.puts '   subgraph cluster_0 {'
  dotFile.puts '     color=white;'
  dotFile.puts '     node [shape="box3d"  style="filled" ];'
  dotFile.puts '     "Workbook" -> "Dashboard" -> "Worksheet" -> "Data Source"'
  dotFile.puts '   }'
end

def closeDot dotFile
  dotFile.puts ' '
  dotFile.puts '// -------------------------------------------------------------'
  dotFile.puts '}'
  dotFile.close
end

def processDashboards twb
  pairs = []
  twb.dashboards.each do |dashName,dash|
    pairs          << [twb.name,dashName]
    sheets = dash.worksheets
    sheets.each do |sheetName,sheet|
      pairs << [dashName,sheetName]
      $worksheets.delete sheetName
      $sheetCnt += 1
    end
    $dashCnt += 1
  end
  return pairs
end

def processWorksheets twb
  pairs = []
  worksheets  = twb.worksheets
  worksheets.each do |sheetName,sheet|
    datasources = sheet.datasources
    datasources.each do |dsName,ds|
      pairs        << [sheetName,ds.uiname]
      $datasources.delete ds.uiname
    end
  end
  return pairs
end

def processOrphans twb
  pairs = []
  $datasources.each { |dsn| pairs << [twb.name,dsn] }
  $worksheets.each  { |wsn| pairs << [twb.name,wsn] }
  return pairs
end

#system 'cls'


init

puts "\n Graphing Workbooks and their Dashboards, Worksheets, and Data Sources.\n\n"

path = if ARGV.empty? then '*.twb' else ARGV[0] end
puts " Looking for Workbooks matching: #{path}"
Dir.glob(path) { |twb| processTWB twb }

plural = if $dashCnt == 1 then '' else 's' end

puts "\n\tDone.\n\tFound #{$sheetCnt} Worksheets in #{$dashCnt} Dashboard#{plural} of #{$twbCnt} Workbooks scanned"

$csv.close unless $csv.nil?
