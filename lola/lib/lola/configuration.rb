# frozen_string_literal: true

module Lola
  class Configuration
    attr_accessor :url, :timeout, :documents_path

    def initialize
      @url = ENV.fetch("LOLA_URL", "http://localhost:8080")
      @timeout = ENV.fetch("LOLA_TIMEOUT", 120).to_i
      @documents_path = ENV.fetch("LOLA_DOCUMENTS_PATH", "/documents")
    end
  end
end
