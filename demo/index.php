<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>FastExcelWriter Demo</title>
</head>
<body>
<?php
$files = glob(__DIR__ . '/demo*.*');

foreach ($files as $file) {
    $name = basename($file);
    echo "<a href=\"$name\" target=\"_blank\">$name</a><br>" ;
}
?>
</body>
</html>