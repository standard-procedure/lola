# frozen_string_literal: true

require_relative "lola/version"
require_relative "lola/configuration"
require_relative "lola/errors"
require_relative "lola/client"

module Lola
  class << self
    attr_writer :configuration

    def configuration
      @configuration ||= Configuration.new
    end

    def configure
      yield(configuration)
    end

    def reset_configuration!
      @configuration = Configuration.new
    end
  end
end
