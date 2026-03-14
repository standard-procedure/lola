# frozen_string_literal: true

require_relative "lib/lola/version"

Gem::Specification.new do |spec|
  spec.name = "lola"
  spec.version = Lola::VERSION
  spec.authors = ["Collabor8Online"]
  spec.email = ["dev@collabor8online.co.uk"]

  spec.summary = "Ruby client for the Lola document processing service"
  spec.description = "Wraps a Python/LibreOffice microservice to execute mail merges " \
                     "against Word .docx templates with native MERGEFIELD codes, " \
                     "convert documents between formats, and extract merge field names."
  spec.homepage = "https://github.com/collabor8online/lola"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 3.1.0"

  spec.files = Dir.chdir(__dir__) do
    `git ls-files -z`.split("\x0").reject do |f|
      (File.expand_path(f) == __FILE__) ||
        f.start_with?("spec/", "test/", ".git", ".github", "Gemfile")
    end
  end

  spec.require_paths = ["lib"]

  spec.add_development_dependency "rspec", "~> 3.13"
  spec.add_development_dependency "webmock", "~> 3.23"
  spec.add_development_dependency "rake", "~> 13.0"
end
