# frozen_string_literal: true

require "net/http"
require "json"
require "uri"

module Lola
  class Client
    def initialize(url: nil, timeout: nil)
      @url = url || Lola.configuration.url
      @timeout = timeout || Lola.configuration.timeout
    end

    # Check service health
    # @return [Hash] status information
    def health
      get("/health")
    end

    # Convert a document to another format
    # @param input_path [String] path to source document (relative to /documents)
    # @param format [Symbol] target format (:pdf, :docx, :odt, :html, :rtf)
    # @param output_path [String, nil] output file path (auto-generated if nil)
    # @return [Hash] conversion result with output_path, format, size_bytes, duration_ms
    def convert(input_path, format: :pdf, output_path: nil)
      post("/convert", {
        input_path: input_path,
        output_format: format.to_s,
        output_path: output_path
      }.compact)
    end

    # Extract merge field names from a template
    # @param template_path [String] path to template (relative to /documents)
    # @return [Array<String>] sorted list of field names
    def fields(template_path)
      response = get("/fields", template_path: template_path)
      response["fields"]
    end

    # Execute a mail merge
    # @param template [String] path to template (relative to /documents)
    # @param data [Array<Hash>] array of data records
    # @param output_dir [String] output directory (relative to /documents)
    # @param output_format [Symbol] output format (:pdf, :docx, :odt)
    # @param filename_field [String, nil] data field for output filenames
    # @return [Hash] merge result with output_files, record_count, duration_ms
    def mail_merge(template:, data:, output_dir:, output_format: :pdf, filename_field: nil)
      post("/mail_merge", {
        template_path: template,
        data: data,
        output_dir: output_dir,
        output_format: output_format.to_s,
        filename_field: filename_field
      }.compact)
    end

    private

    def get(path, params = {})
      uri = build_uri(path, params)
      request = Net::HTTP::Get.new(uri)
      execute(request, uri)
    end

    def post(path, body)
      uri = build_uri(path)
      request = Net::HTTP::Post.new(uri)
      request.content_type = "application/json"
      request.body = JSON.generate(body)
      execute(request, uri)
    end

    def execute(request, uri)
      http = Net::HTTP.new(uri.host, uri.port)
      http.use_ssl = uri.scheme == "https"
      http.read_timeout = @timeout
      http.open_timeout = 10
      response = http.request(request)
      handle_response(response)
    rescue Errno::ECONNREFUSED, Errno::ECONNRESET, Errno::EHOSTUNREACH => e
      raise Lola::ConnectionError, "Cannot connect to Lola service at #{@url}: #{e.message}"
    rescue Net::ReadTimeout, Net::OpenTimeout => e
      raise Lola::TimeoutError, "Request to Lola service timed out: #{e.message}"
    end

    def handle_response(response)
      body = JSON.parse(response.body)

      case response.code.to_i
      when 200..299
        body
      when 400..499
        raise error_for_code(body["code"], body["error"])
      when 500..599
        raise Lola::ServiceError.new(body["error"] || "Internal server error", code: body["code"])
      else
        raise Lola::Error, "Unexpected response: #{response.code}"
      end
    rescue JSON::ParserError
      raise Lola::Error, "Invalid response from Lola service: #{response.body}"
    end

    def error_for_code(code, message)
      case code
      when "CONVERSION_ERROR"   then Lola::ConversionError.new(message, code: code)
      when "MERGE_ERROR"        then Lola::MergeError.new(message, code: code)
      when "TEMPLATE_ERROR"     then Lola::TemplateError.new(message, code: code)
      when "TEMPLATE_NOT_FOUND" then Lola::TemplateError.new(message, code: code)
      when "FIELD_NOT_FOUND"    then Lola::FieldNotFoundError.new(message, code: code)
      when "MISSING_FIELDS"     then Lola::FieldNotFoundError.new(message, code: code)
      when "INVALID_FORMAT"     then Lola::ConversionError.new(message, code: code)
      when "INVALID_REQUEST"    then Lola::Error.new(message)
      when "LIBREOFFICE_ERROR"  then Lola::ServiceError.new(message, code: code)
      when "TIMEOUT"            then Lola::TimeoutError.new(message)
      else Lola::ServiceError.new(message || "Unknown error", code: code)
      end
    end

    def build_uri(path, params = {})
      uri = URI.join(@url, path)
      uri.query = URI.encode_www_form(params) unless params.empty?
      uri
    end
  end
end
