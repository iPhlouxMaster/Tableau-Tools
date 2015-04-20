# X-Ray Dashboards.rb - this Ruby script Copyright 2013-2015 Christopher Gerrard

require 'twb'

system "cls"

$doctwb      = true
$replacetwb  = false
$dashdoclbl  = 'dashdoc'
$localurl    = 'file:///' + Dir.pwd + '/'

  def xray twbname
    twb    = Twb::Workbook.new(twbname)
    xrayer = Twb::DashboardXRayer.new(twb)
    xrays  = xrayer.xray
    cnt    = 0
    xrays.each do |dash, html|
      htmlfilename =  twb.name + '.' + dash.to_s  + '.html'
      saveHTML(htmlfilename, html)
      cnt += 1
      if $doctwb
        inject(twb, dash.to_s, htmlfilename)
      end
    end
    puts "\t    #{cnt} \t       #{twbname}"
  end
  
  def saveHTML(htmlfilename, html)
    begin
      htmlfile = File.open(htmlfilename, 'w')
      htmlfile.puts html
      htmlfile.close
    rescue
      # Common failure is when the Dashboard name contains
      # invalid file name Characters. or when the name is
      # an invalid file name.
      # Stripping the non-ASCII characters from the Dashboard
      # name fixes this, in the cases see so far.
      # This rescue-recursion technique can potentially cause
      # an infite-loop condition. (not seen, but possible)
      saveHTML( sanitize(htmlfilename), html)
    end
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
  
  def sanitize(str)
    str.gsub(/[^a-z0-9\-]+/i, ' ')
  end

system 'cls'
puts "\n\n\t X-raying Dashboards\n\n\t # Dashboards  Workbook\n\t ------------  ----------------------"

path = if ARGV.empty? then '*.twb' else ARGV[0] end
Dir.glob(path) {|twb| xray twb }
