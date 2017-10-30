require 'creek'
require 'time'
require 'google_calendar'
require 'yaml'
require 'open-uri'

conf = YAML.load(File.read("sparta.yml"))
xlx_file = open('http://www.hcsparta.cz/ledy/2017_2018_LEDOVE_PLOCHY.xlsx')
creek = Creek::Book.new xlx_file, check_file_extension: false
@cal = Google::Calendar.new(:client_id     => conf['client'],
                           :client_secret => conf['secret'],
                           :calendar      => conf['calendar'],
                           :refresh_token => conf['token'],
                           :timezone      => 'Europe/Prague',
                           :redirect_url  => "urn:ietf:wg:oauth:2.0:oob" # this is what Google uses for 'applications'
                           )
@times = {}
time = Time.new('2017','10','27','06','00')
dates = Hash.new
results = Array.new
books = creek.sheets
index = "C"
count = 1
row_count = 1
@lookup_entry = conf['entry']

while count <= 80
  @times[index] = time
  index = index.next
  time += 900
  count += 1
end

def calendar_cleanup
  puts "Cleaning up"
  events = @cal.find_events("#{@lookup_entry}")
  events.each do |event|
    event.delete
  end
end

def get_index(index)
  return index.split(%r{[[:digit:]]}).first
end

def get_line(index)
  return index.split(%r{[[:alpha:]]}).last.to_i
end

def get_offset(hash={},line,index)
  counter = 1
  until !hash.fetch("#{index}#{line}").nil?
    counter += 1
    index = index.next
  end
  return counter
end

def get_date(num)
  return num + 2
end

def gen_ical(hash={})
   date = hash.fetch("date")
   start = Time.parse("#{date} #{hash.fetch('start')}")
   dend = Time.parse("#{date} #{hash.fetch('end')}")
   @cal.create_event do |e|
     e.description = hash.fetch("event")
     e.start_time = start
     e.title = "SPARTA: #{hash.fetch("event")}"
     e.end_time = dend
     e.location = "#{hash.fetch("location")}"
   end
end
def get_arena(arena)
  return 'TIPSPORT Arena' if arena.start_with?("TIPSPORT")
  return 'Mala Sportovni Hala' if arena.start_with?("MAL")
end
calendar_cleanup
books.each do |sheet|
  arena = get_arena(sheet.name)
  sheet.rows.each do |row|
    if row.has_key?("A#{row_count}") 
      date = { "A#{row_count}" => row.fetch("A#{row_count}")}
      dates.merge!(date)
    end
    row.each do |k,v|
      if v == @lookup_entry
        line = get_line(k)
        index = get_index(k)
        startt = @times.fetch(index)
        offset = get_offset(row, line, index.next) * 15
        endt = (startt + (offset * 60 )).strftime("%H:%M:%S")
        result = { "event" => v, "location" => arena, "line" => line, "start" => startt.strftime('%H:%M:%S'), "end" => endt,"date" => "" }
        results.push(result)
      end
    end
    row_count += 1
  end
  results.each do |result|
    entry = get_date(result.fetch("line"))
    if result.fetch("date").empty?
      result.merge!({ "date" => dates.fetch("A#{entry}").strftime('%y/%m/%d') })
      gen_ical(result)
    end
  end
end
