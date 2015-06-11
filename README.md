# Bayonet gem

[![Build Status](https://travis-ci.org/nethad/bayonet.svg?branch=master)](https://travis-ci.org/nethad/bayonet)
[![Code Climate](https://codeclimate.com/github/nethad/bayonet/badges/gpa.svg)](https://codeclimate.com/github/nethad/bayonet)
[![Test Coverage](https://codeclimate.com/github/nethad/bayonet/badges/coverage.svg)](https://codeclimate.com/github/nethad/bayonet/coverage)
[![Gem Version](https://badge.fury.io/rb/bayonet.svg)](http://badge.fury.io/rb/bayonet)

Bayonet is a Microsoft Excel write-only gem that reads and produces XLSX files.
It's strength lies in the fact that it's able to open bigger Excel files (even with macros!) -- and write cells without touching the rest of the Excel file. I've written the gem because *roo* was unbearably slow and was using ~1GB of RAM to read a 5MB XLSX file.


**WARNING:** Use this gem at your own risk. It's writing directly to sheet XML files and uses a few tricks that Microsoft Excel 2010 seems to be fine with, but I cannot guarantee that future (or past) versions will play along. That being said, I've used it successfully in one of my projects. If you find any bugs, let me know.

### Usage example

```ruby
workbook = Bayonet::Workbook.new(a_file_path)

workbook.on_sheet('The First Sheet') do |sheet|
  sheet.set_typed_cell('A', 1, "I'm a string") # set a cell auto-typed to a string
  sheet.set_typed_cell('B', 1, 42)             # set a cell auto-typed to a number
  sheet.set_cell('C', 1, "Some value")         # set a cell without setting its type
  sheet.write_string('D', 1, "Some value")     # set a cell and forcing it to be a string
  sheet.write_number('E', 1, 23)               # set a cell and forcing it to be a number
end

workbook.write_and_close(output_file_path)
```

See the *Tips* section for more information.

### When to use Bayonet?

* If you want to modify an existing, possibly huge XLSX file.
* If other gems fail at the task or are too slow.

### When should I use other gems?

* If you want to create an XLSX file from scratch.
* If you plan to open legacy (binary) XLS files.
* If your XLSX file is small and other gems are working fine.

Other gems I can recommend are:

* [roo](https://rubygems.org/gems/roo) (read/write, lots of features and supported formats)
* [axlsx](https://rubygems.org/gems/axlsx) (write-only)
* [creek](https://rubygems.org/gems/creek) (read-only, very fast)


### Tips

#### When should I use `set_typed_cell` vs. `set_cell`?

Normally you should use `set_typed_cell`, which sets the cell type correctly and falls back to `string`. If you're using a prepared XLSX sheet you might not want to override the current cell type. You should use `set_cell` for those instances. If you want to force a cell type, you can use `write_string` or `write_number` instead, but this should rarely be necessary.

#### Why can't I read cells?

It's simply not been implemented yet. I've been writing to template files with pre-defined cells to fill, so reading cells was not necessary. It shouldn't be hard to implement, though there might be some edge cases I haven't considered yet, that's why I left it out for now. Pull requests are welcome, otherwise I'd suggest to use other gems for that.

