<?php
include 'db.php';

$id = $_GET['id'];
$conn->query("DELETE FROM productos WHERE id = $id");

header('Location: listar_productos.php');
exit;
?>
