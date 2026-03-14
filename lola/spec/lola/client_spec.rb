# frozen_string_literal: true

require "spec_helper"

RSpec.describe Lola::Client do
  subject(:client) { described_class.new(url: "http://localhost:8080") }

  describe "#health" do
    it "returns service status" do
      stub_request(:get, "http://localhost:8080/health")
        .to_return(
          status: 200,
          body: '{"status":"ok","libreoffice":"connected","version":"0.1.0"}',
          headers: {"Content-Type" => "application/json"}
        )

      result = client.health
      expect(result["status"]).to eq("ok")
      expect(result["libreoffice"]).to eq("connected")
    end
  end

  describe "#convert" do
    it "converts a document" do
      stub_request(:post, "http://localhost:8080/convert")
        .with(body: {input_path: "templates/invoice.docx", output_format: "pdf"})
        .to_return(
          status: 200,
          body: '{"output_path":"output/invoice.pdf","format":"pdf","size_bytes":12345,"duration_ms":2000}',
          headers: {"Content-Type" => "application/json"}
        )

      result = client.convert("templates/invoice.docx", format: :pdf)
      expect(result["output_path"]).to eq("output/invoice.pdf")
    end

    it "raises ConversionError on failure" do
      stub_request(:post, "http://localhost:8080/convert")
        .to_return(
          status: 422,
          body: '{"error":"Conversion failed","code":"CONVERSION_ERROR"}',
          headers: {"Content-Type" => "application/json"}
        )

      expect { client.convert("bad.docx") }.to raise_error(Lola::ConversionError)
    end
  end

  describe "#fields" do
    it "returns merge field names" do
      stub_request(:get, "http://localhost:8080/fields?template_path=templates/letter.docx")
        .to_return(
          status: 200,
          body: '{"template_path":"templates/letter.docx","fields":["Address","City","CustomerName"],"field_count":3}',
          headers: {"Content-Type" => "application/json"}
        )

      fields = client.fields("templates/letter.docx")
      expect(fields).to eq(["Address", "City", "CustomerName"])
    end
  end

  describe "#mail_merge" do
    it "executes a mail merge" do
      stub_request(:post, "http://localhost:8080/mail_merge")
        .to_return(
          status: 200,
          body: '{"output_files":["output/0001.pdf"],"record_count":1,"duration_ms":3000,"warnings":[]}',
          headers: {"Content-Type" => "application/json"}
        )

      result = client.mail_merge(
        template: "templates/letter.docx",
        data: [{"CustomerName" => "Alice"}],
        output_dir: "output",
        output_format: :pdf
      )
      expect(result["record_count"]).to eq(1)
    end
  end

  describe "error handling" do
    it "raises ConnectionError when service is unavailable" do
      stub_request(:get, "http://localhost:8080/health")
        .to_raise(Errno::ECONNREFUSED)

      expect { client.health }.to raise_error(Lola::ConnectionError)
    end

    it "raises TimeoutError on timeout" do
      stub_request(:get, "http://localhost:8080/health")
        .to_raise(Net::ReadTimeout)

      expect { client.health }.to raise_error(Lola::TimeoutError)
    end

    it "raises TemplateError for missing templates" do
      stub_request(:get, "http://localhost:8080/fields?template_path=missing.docx")
        .to_return(
          status: 404,
          body: '{"error":"File not found","code":"TEMPLATE_NOT_FOUND"}',
          headers: {"Content-Type" => "application/json"}
        )

      expect { client.fields("missing.docx") }.to raise_error(Lola::TemplateError)
    end
  end
end
