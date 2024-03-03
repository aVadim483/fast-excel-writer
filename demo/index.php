<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>FastExcelWriter Demo</title>
</head>
<body>

<?php
$files = glob(__DIR__ . '/demo*.*');
usort($files, function ($a, $b) {
    preg_match('/-(\d+)-([a-z0-9\-]+)\./', $a, $m1);
    preg_match('/-(\d+)-([a-z0-9\-]+)\./', $b, $m2);
    if ($m1[1] !== $m2[1]) {
        return ($m1[1] < $m2[1]) ? -1 : 1;
    }

    if ($m1[2] === $m2[2]) {
        return 0;
    }
    return ($m1[2] < $m2[2]) ? -1 : 1;
});

$list1 = $list2 = [];
foreach ($files as $file) {
    $name = basename($file);
    if (preg_match('/-(\d+)-([a-z0-9\-]+)\./', $name, $m)) {
        $text = ucwords(str_replace('-', ' ', $m[2]));
    }
    else {
        $text = $name;
    }
    if (strpos($text, 'Chart ') === 0) {
        $list2[] = [
            'link' => $name,
            'text' => $text,
        ];
    }
    else {
        $list1[] = [
            'link' => $name,
            'text' => $text,
        ];
    }
}
echo '<ul>';
foreach ($list1 as $item) {
    echo "<li><a href=\"{$item['link']}\" target=\"_blank\">{$item['text']}</a></li>" ;
}
echo '</ul><p></p><ul>';
foreach ($list2 as $item) {
    echo "<li><a href=\"{$item['link']}\" target=\"_blank\">{$item['text']}</a></li>" ;
}
echo '</ul>';

?>

</body>
</html>