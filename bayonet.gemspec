# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'bayonet/version'

Gem::Specification.new do |spec|
  spec.name          = "bayonet"
  spec.version       = Bayonet::VERSION
  spec.authors       = ["Thomas Ritter"]
  spec.email         = ["ritter.thomas@gmail.com"]

  spec.summary       = %q{Bayonet is a Microsoft Excel write-only gem that reads and produces XLSX files.}
  spec.description   = %q{Bayonet is a Microsoft Excel write-only gem that reads and produces XLSX files. It's strength lies in the fact that it's able to open bigger Excel files (even with macros!) -- and write cells without touching the rest of the Excel file.}
  spec.homepage      = "https://github.com/nethad/bayonet"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.7"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency "rubyzip", "~> 2.2"
  spec.add_development_dependency "nokogiri", "~> 1.6"
end
