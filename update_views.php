<?php
$servername = "localhost";
$username = "username";
$password = "password";
$dbname = "database";

// 创建连接
$conn = new mysqli($servername, $username, $password, $dbname);

// 检查连接
if ($conn->connect_error) {
    die("连接失败: " . $conn->connect_error);
}

$page = $_POST['page'];
$sql = "UPDATE page_views SET views = views + 1 WHERE page = '$page'";
$conn->query($sql);

$conn->close();
?>