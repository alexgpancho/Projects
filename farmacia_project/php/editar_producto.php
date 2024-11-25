<?php
include 'db.php';

$id = $_GET['id'];
$result = $conn->query("SELECT * FROM productos WHERE id = $id");
$product = $result->fetch_assoc();

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $nombre = $_POST['nombre'];
    $descripcion = $_POST['descripcion'];
    $precio = $_POST['precio'];
    $stock = $_POST['stock'];

    $sql = "UPDATE productos SET nombre = '$nombre', descripcion = '$descripcion', precio = $precio, stock = $stock WHERE id = $id";
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
    <title>Editar Producto</title>
</head>
<body>
    <h1>Editar Producto</h1>
    <form method="POST">
        <label for="nombre">Nombre:</label>
        <input type="text" name="nombre" value="<?= $product['nombre'] ?>" required>
        <br>
        <label for="descripcion">Descripción:</label>
        <textarea name="descripcion" required><?= $product['descripcion'] ?></textarea>
        <br>
        <label for="precio">Precio:</label>
        <input type="number" step="0.01" name="precio" value="<?= $product['precio'] ?>" required>
        <br>
        <label for="stock">Stock:</label>
        <input type="number" name="stock" value="<?= $product['stock'] ?>" required>
        <br>
        <button type="submit">Actualizar</button>
    </form>
</body>
</html>
