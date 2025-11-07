## FastExcelWriter - Notes

There are currently two types of comments in Excel - **comments** and **notes**
(see [The difference between threaded comments and notes](https://support.microsoft.com/en-us/office/the-difference-between-threaded-comments-and-notes-75a51eec-4092-42ab-abf8-7669077b7be3)).
Notes are old style comments in Excel (text on a light yellow background).
You can add notes to any cells using method ```addNote()```

```php

$sheet->writeCell('Text to A1');
$sheet->addNote('A1', 'This is a note for cell A1');

$sheet->writeCell('Text to B1')->addNote('This is a note for B1');
$sheet->writeTo('C4', 'Text to C4')->addNote('Note for C1');

// If you specify a range of cells, then the note will be added to the left top cell
$sheet->addNote('E4:F8', "This note\nwill added to E4");

// You can split text into multiple lines
$sheet->addNote('D7', "Line 1\nLine 2");

```

You can change some note options. Allowed options of a note are:

* **width** - default value is ```'96pt'```
* **height** - default value is ```'55.5pt'```
* **fill_color** - default value is ```'#FFFFE1'```
* **show** - default value is ```false```

```php

$sheet->addNote('A1', 'This is a note for cell A1', 
    ['width' => '200pt', 'height' => '100pt', 'fill_color' => '#ffcccc']);

// Parameters "width" and "height" can be numeric, by default these values are in points
// The "fill_color" parameter can be shortened
$noteStyle = [
    'width' => 200, // equivalent to '200pt'
    'height' => 100, // equivalent to '100pt'
    'fill_color' => 'fcc', // equivalent to '#ffcccc'
];
$sheet->writeCell('Text to B1')->addNote('This is a note for B1', $noteStyle);

// This note is visible when the Excel workbook is displayed
$sheet->addNote('C8', 'This note is always visible', ['show' => true]);
```

Also, you can use rich text in notes

```php
$richText = new \avadim\FastExcelWriter\RichText('here is <c=f00>red</c> and <c=00f>blue</c> text');
$sheet->addNote('C8', $richText);
```

For more information on using rich text, see here: [Using Rich Text](/docs/03-writing.md#using-rich-text)
