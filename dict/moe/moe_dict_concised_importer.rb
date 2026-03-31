#!/usr/bin/env ruby
# frozen_string_literal: true

# Imports 教育部國語辭典簡編本 from the MoE download page and converts it to
# dict/moe/moe_dict_concised.csv for use with libchewing.
#
# Usage:
#   ruby dict/moe/moe_dict_concised_importer.rb              # download from MoE
#   ruby dict/moe/moe_dict_concised_importer.rb path/to.zip  # use local ZIP
#   ruby dict/moe/moe_dict_concised_importer.rb path/to.xlsx # use local XLSX

require 'net/http'
require 'uri'
require 'tmpdir'
require 'fileutils'

DOWNLOAD_PAGE = 'https://language.moe.gov.tw/001/Upload/Files/site_content/M0001/respub/dict_concised_download.html'
SOURCES_DIR   = File.expand_path('sources', __dir__)
OUTPUT_FILE   = File.expand_path('moe_dict_concised.csv', __dir__)

DC_TITLE   = '教育部國語辭典簡編本詞庫'
DC_RIGHTS  = '中華民國教育部（Ministry of Education, R.O.C.）。《國語辭典簡編本》（版本編號：%s ）網址： http://dict.concised.moe.edu.tw /'
DC_LICENSE = 'CC-BY-ND-3.0-TW'

# ── HTTP helpers ─────────────────────────────────────────────────────────────

def http_get(url, limit: 5)
  raise 'Too many redirects' if limit.zero?

  uri = URI(url)
  response = Net::HTTP.start(uri.host, uri.port, use_ssl: uri.scheme == 'https',
                             open_timeout: 30, read_timeout: 120) do |http|
    http.get(uri.request_uri, 'User-Agent' => 'Mozilla/5.0')
  end

  case response
  when Net::HTTPSuccess     then response
  when Net::HTTPRedirection then http_get(response['location'], limit: limit - 1)
  else raise "HTTP #{response.code}: #{url}"
  end
end

def download_file(url, dest)
  uri = URI(url)
  puts "  GET #{url}"
  Net::HTTP.start(uri.host, uri.port, use_ssl: uri.scheme == 'https',
                  open_timeout: 30, read_timeout: 300) do |http|
    http.request(Net::HTTP::Get.new(uri, 'User-Agent' => 'Mozilla/5.0')) do |response|
      raise "HTTP #{response.code}" unless response.is_a?(Net::HTTPSuccess)

      File.open(dest, 'wb') { |f| response.read_body { |chunk| f.write(chunk) } }
    end
  end
end

# ── URL / version helpers ─────────────────────────────────────────────────────

# Find the text-database ZIP link on the download page (exclude pic/music links).
def find_zip_url(html, base_url)
  base = base_url.sub(%r{[^/]+\z}, '')
  href = html.scan(/href=["']?([^"'\s>]*dict_concised_\d{4}_\d{8}\.zip)["']?/)
              .map(&:first)
              .reject { |h| h.include?('pic') || h.include?('music') }
              .first
  raise 'Could not find text-database ZIP link on download page' unless href

  URI.join(base, href).to_s
end

# Extract the MoE version string (e.g. "2014_20251229") from a filename/URL.
def version_from_name(name)
  name[/dict_concised_(\d{4}_\d{8})/, 1]
end

# Format MoE version "2014_YYYYMMDD" → dc:identifier "YYYY.MM.DD"
def identifier_from_version(version)
  date = version.split('_').last
  "#{date[0, 4]}.#{date[4, 2]}.#{date[6, 2]}"
end

# ── Fast XLSX parsers (regex-based, no gems required) ─────────────────────────

XML_ENTITIES = { '&amp;' => '&', '&lt;' => '<', '&gt;' => '>', '&quot;' => '"', '&apos;' => "'" }.freeze

def unescape_xml(str)
  str.gsub(/&(?:amp|lt|gt|quot|apos);/, XML_ENTITIES)
end

# Read xl/sharedStrings.xml and return a flat array of strings.
def parse_shared_strings(path)
  content = IO.popen(['unzip', '-p', path, 'xl/sharedStrings.xml'], encoding: 'UTF-8', &:read)
  content.scan(/<si>(.*?)<\/si>/m).map do |si_body,|
    unescape_xml(si_body.scan(/<t[^>]*>(.*?)<\/t>/m).map { |t,| t }.join)
  end
end

# Read xl/worksheets/sheet1.xml and return [[phrase, bopomofo], ...] entries.
# We only need columns A (phrase), G (main bopomofo), I (variant bopomofo).
PHRASE_COL       = 'A'
BOPOMOFO_COL     = 'G'
VARIANT_BOPO_COL = 'I'
WANTED_COLS      = [PHRASE_COL, BOPOMOFO_COL, VARIANT_BOPO_COL].freeze

def normalize_bopomofo(str)
  str.to_s.gsub("\u3000", ' ').strip
     .gsub(/˙([ㄅ-ㄩ]+)/, '\1˙')
     .gsub(/([ˊˇˋ˙])([ㄅ-ㄩ])/, '\1 \2')
     .gsub(/([ㄅ-ㄩ])ㄦ/, '\1 ㄦ')
end

def parse_sheet_entries(xlsx_path, strings)
  content = IO.popen(['unzip', '-p', xlsx_path, 'xl/worksheets/sheet1.xml'], encoding: 'UTF-8', &:read)

  entries  = []
  row_num  = 0

  content.scan(/<row\b[^>]*>(.*?)<\/row>/m) do |row_body,|
    row_num += 1
    next if row_num == 1  # first row is the header

    cells = {}
    row_body.scan(/<c\s([^>]*)>(.*?)<\/c>/m) do |attrs, body|
      col_m  = attrs.match(/\br="([A-Z]+)\d+"/)
      next unless col_m

      col = col_m[1]
      next unless WANTED_COLS.include?(col)

      v_m = body.match(/<v>(.*?)<\/v>/m)
      next unless v_m

      type  = attrs.match(/\bt="([^"]+)"/)&.[](1)
      raw   = v_m[1]
      cells[col] = type == 's' ? strings[raw.to_i].to_s : raw
    end

    phrase = cells[PHRASE_COL].to_s.strip.gsub(/[，、。！？；：「」『』【】〔〕《》〈〉…—～·]/, '')
    bopo   = normalize_bopomofo(cells[BOPOMOFO_COL])
    next if phrase.empty? || bopo.empty?

    entries << [phrase, bopo]

    variant = normalize_bopomofo(cells[VARIANT_BOPO_COL])
    entries << [phrase, variant] if variant.length.positive? && variant != bopo
  end

  entries
end

# ── CSV output ────────────────────────────────────────────────────────────────

def write_csv(entries, version)
  identifier = identifier_from_version(version)
  File.open(OUTPUT_FILE, 'w', encoding: 'UTF-8') do |f|
    f.puts "# dc:title,#{DC_TITLE},"
    f.puts "# dc:rights,#{format(DC_RIGHTS, version)},"
    f.puts "# dc:license,#{DC_LICENSE},"
    f.puts "# dc:identifier,#{identifier},"
    entries.each { |phrase, bopo| f.puts "#{phrase},1,#{bopo}" }
  end
end

# ── Main ──────────────────────────────────────────────────────────────────────

local_input = ARGV[0]

Dir.mktmpdir do |tmpdir|
  xlsx_path = nil
  version   = nil

  if local_input&.end_with?('.xlsx')
    xlsx_path = File.expand_path(local_input)
    version   = version_from_name(File.basename(xlsx_path)) || 'unknown'
  elsif local_input&.end_with?('.zip')
    zip_path = File.expand_path(local_input)
    puts "Extracting #{zip_path}..."
    system('unzip', '-q', '-o', zip_path, '-d', tmpdir) or raise 'unzip failed'
    xlsx_path = Dir.glob("#{tmpdir}/*.xlsx").first
    raise 'XLSX not found in ZIP' unless xlsx_path
    version = version_from_name(File.basename(xlsx_path)) ||
              version_from_name(File.basename(zip_path)) ||
              'unknown'
  else
    puts 'Fetching download page...'
    html    = http_get(DOWNLOAD_PAGE).body.force_encoding('UTF-8')
    zip_url = find_zip_url(html, DOWNLOAD_PAGE)
    version = version_from_name(zip_url)
    puts "Found ZIP: #{zip_url} (version #{version})"

    FileUtils.mkdir_p(SOURCES_DIR)
    zip_path = File.join(SOURCES_DIR, File.basename(zip_url))
    puts 'Downloading...'
    download_file(zip_url, zip_path)

    puts 'Extracting ZIP...'
    system('unzip', '-q', '-o', zip_path, '-d', tmpdir) or raise 'unzip failed'
    xlsx_path = Dir.glob("#{tmpdir}/*.xlsx").first
    raise 'XLSX not found after extraction' unless xlsx_path
  end

  puts "Parsing #{File.basename(xlsx_path)}..."
  strings = parse_shared_strings(xlsx_path)
  entries = parse_sheet_entries(xlsx_path, strings)
  puts "Parsed #{entries.size} entries"

  puts "Writing #{OUTPUT_FILE}..."
  write_csv(entries, version)
end

puts 'Done.'
