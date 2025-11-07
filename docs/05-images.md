## FastExcelWriter - Images

You can insert image to sheet from local file, URL or image string in base64

```php
$sheet->addImage($cell, $imageFile, $imageStyle);

// Insert an image to the cell A1 from local path
$sheet->addImage('A1', 'path/to/file');

// Insert an image to the cell A1 from URL
$sheet->addImage('A1', 'https://site.com/image.jpg');

// Insert an image to the cell A1 from base64 string
$sheet->addImage('A1', 'data:image/jpeg;base64,/9j/4AAQ...');

// Insert an image to the cell B2 and set with to 150 pixels (height will change proportionally)
$sheet->addImage('B2', 'path/to/file', ['width' => 150]);

// Set height to 150 pixels (with will change proportionally)
$sheet->addImage('C3', 'path/to/file', ['height' => 150]);

// Set size in pixels
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150]);

// Add hyperlink to the image
$sheet->addImage('D4', 'path/to/file', ['width' => 150, 'height' => 150, 'hyperlink' => 'https://www.google.com/']);
```

Available keys of image style:

* 'width' -- width of image
* 'height' -- height of image
* 'hyperlink' -- URL of hyperlink
* 'x' -- offset in pixels relative to the left border of the cell
* 'y' -- offset in pixels relative to the top border of the cell

**IMPORTANT:** in MS Excel, value 'x' cannot be greater than the column width of the parent cell,
and value 'y' cannot be greater than the row height
