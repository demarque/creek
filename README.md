# Creek - Stream parser for large Excel (xlsx and xlsm) files.

Creek is a Ruby gem that provides a fast, simple and efficient method of parsing large Excel (xlsx and xlsm) files.


## Installation

Creek can be used from the command line or as part of a Ruby web framework. To install the gem using terminal, run the following command:

```
gem install creek
```

To use it in Rails, add this line to your Gemfile:

```ruby
gem 'creek'
```

## Basic Usage
Creek can simply parse an Excel file by looping through the rows enumerator:

```ruby
require 'creek'
creek = Creek::Book.new 'spec/fixtures/sample.xlsx'
sheet = creek.sheets[0]

sheet.rows.each do |row|
  puts row # => ["Content 1", nil, nil, "Content 3"]
end

sheet.rows(headers: true).each do |row|
  puts row # => { 'header1' => "Content 1", 'header2' => nil, 'header3' => nil, 'header4' => "Content 3" }
end

sheet.rows_with_meta_data.each do |row|
  puts row # => {"collapsed"=>"false", "customFormat"=>"false", "customHeight"=>"true", "hidden"=>"false", "ht"=>"12.1", "outlineLevel"=>"0", "r"=>"1", "header_row" => false, "cells"=>["Content 1", nil, nil, "Content 3"]}
end

sheet.rows_with_meta_data(headers: true).each do |row|
  puts row # => {"collapsed"=>"false", "customFormat"=>"false", "customHeight"=>"true", "hidden"=>"false", "ht"=>"12.1", "outlineLevel"=>"0", "r"=>"1", "header_row" => false, "cells"=>{ 'header1' => "Content 1", 'header2' => nil, 'header3' => nil, 'header4' => "Content 3" }}
end

sheet.state   # => 'visible'
sheet.name    # => 'Sheet1'
sheet.rid     # => 'rId2'
```

## Headers
`rows` and `rows_with_meta_data` both accept the kwargs `headers` and `header_row_number` to load
the rows as hash with the headers as keys. Also, a `header_row` boolean is added to the row metadata.
See examples above.

Headers (as an array) are loaded once by parsing the file a first time until the `header_row_number` is reached.
Rows are then returned normally as an Enumerator as usual (new Enumerator instance starting from the beginning of the file, that will include header row as well). It's the caller's responsibility to filter the header row as needed.

`extract_headers` can also be called manually from the sheet instance.
Once extracted, the headers can be accessed through the `headers` attr_reader.
As headers are matched to their respective value in the row by index, it's possible to modifies the array in `headers` to customize the headers (to fix typo, make them unique, etc.).

```ruby
creek = Creek::Book.new 'spec/fixtures/sample.xlsx'
sheet = creek.sheets[0]

# Parse the file up to row 3 (file starts at row 1)
sheet.extract_headers(3)
# => ['Header1', 'Other Header', 'More header']

sheet.headers
# => ['Header1', 'Other Header', 'More header']

# Headers can be modified before parsing the file to customize them
sheet.headers[0] = 'A better Header'
sheet.headers
# => ['A better Header', 'Other Header', 'More header']

# Parse the rows as hashes, including the (modified) headers
sheet.rows(headers: true).each do |row|
  puts row # => { 'A better Header' => "Content 1", 'Other Header' => nil, 'More header' => nil }
end

# Or both can be done directly when accessing rows
sheet2 = creek.sheets[1]
sheet2.rows(headers: true, header_row_number: 3).each do |row|
  puts row # => { 'Header1' => "Content 2", 'Other Header' => nil, 'More header' => nil }
end
```

## Filename considerations
By default, Creek will ensure that the file extension is either *.xlsx or *.xlsm, but this check can be circumvented as needed:

```ruby
path = 'sample-as-zip.zip'
Creek::Book.new path, :check_file_extension => false
```

By default, the Rails [file_field_tag](http://api.rubyonrails.org/classes/ActionView/Helpers/FormTagHelper.html#method-i-file_field_tag) uploads to a temporary location and stores the original filename with the StringIO object. (See [this section](http://guides.rubyonrails.org/form_helpers.html#uploading-files) of the Rails Guides for more information.)

Creek can parse this directly without the need for file upload gems such as Carrierwave or Paperclip by passing the original filename as an option:

```ruby
# Import endpoint in Rails controller
def import
  file = params[:file]
  Creek::Book.new file.path, check_file_extension: false
end
```

## Parsing images
Creek does not parse images by default. If you want to parse the images,
use `with_images` method before iterating over rows to preload images information. If you don't call this method, Creek will not return images anywhere.

Cells with images will be an array of Pathname objects.
If an image is spread across multiple cells, same Pathname object will be returned for each cell.

```ruby
sheet.with_images.rows.each do |row|
  puts row # => {"A1"=>[#<Pathname:/var/folders/ck/l64nmm3d4k75pvxr03ndk1tm0000gn/T/creek__drawing20161101-53599-274q0vimage1.jpeg>], "B2"=>"Fluffy"}
end
```

Images for a specific cell can be obtained with images_at method:

```ruby
puts sheet.images_at('A1') # => [#<Pathname:/var/folders/ck/l64nmm3d4k75pvxr03ndk1tm0000gn/T/creek__drawing20161101-53599-274q0vimage1.jpeg>]

# no images in a cell
puts sheet.images_at('C1') # => nil
```

Creek will most likely return nil for a cell with images if there is no other text cell in that row - you can use *images_at* method for retrieving images in that cell.

## Contributing

Contributions are welcomed. You can fork a repository, add your code changes to the forked branch, ensure all existing unit tests pass, create new unit tests which cover your new changes and finally create a pull request.

After forking and then cloning the repository locally, install the Bundler and then use it
to install the development gem dependencies:

```
gem install bundler
bundle install
```

Once this is complete, you should be able to run the test suite:

```
rake
```

## Bug Reporting

Please use the [Issues](https://github.com/pythonicrubyist/creek/issues) page to report bugs or suggest new enhancements.


## License

Creek has been published under [MIT License](https://github.com/pythonicrubyist/creek/blob/master/LICENSE.txt)
