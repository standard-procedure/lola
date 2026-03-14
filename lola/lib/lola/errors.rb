# frozen_string_literal: true

module Lola
  # Base error class for all Lola errors
  class Error < StandardError; end

  # Connection/HTTP errors
  class ConnectionError < Error; end
  class TimeoutError < Error; end

  # Service errors (returned by the Lola service)
  class ServiceError < Error
    attr_reader :code

    def initialize(message, code: nil)
      @code = code
      super(message)
    end
  end

  # Document conversion failed
  class ConversionError < ServiceError; end

  # Mail merge execution failed
  class MergeError < ServiceError; end

  # Template file is invalid or corrupt
  class TemplateError < ServiceError; end

  # Template merge field not found in data
  class FieldNotFoundError < MergeError; end
end
