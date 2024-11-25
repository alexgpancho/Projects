<?php
include 'db.php';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $nombre = $_POST['nombre'];
    $descripcion = $_POST['descripcion'];
    $precio = $_POST['precio'];
    $stock = $_POST['stock'];

    $sql = "INSERT INTO productos (nombre, descripcion, precio, stock) VALUES ('$nombre', '$descripcion', $precio, $stock)";
    $conn->query($sql);

    header('Location: listar_productos.php');
    exit;
}
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Crear Producto</title>
</head>
<body>
    <h1>Crear Producto</h1>
    <form method="POST">
        <label for="nombre">Nombre:</label>
        <input type="text" name="nombre" required>
        <br>
        <label for="descripcion">Descripción:</label>
        <textarea name="descripcion" required></textarea>
        <br>
        <label for="precio">Precio:</label>
        <input type="number" step="0.01" name="precio" required>
        <br>
        <label for="stock">Stock:</label>
        <input type="number" name="stock" required>
        <br>
        <button type="submit">Crear</button>
    </form>
</body>
</html>
