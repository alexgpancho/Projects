<?php
include 'db.php';

$result = $conn->query("SELECT * FROM productos");
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Gestión de Productos</title>
    <link rel="stylesheet" href="../css/styles.css">
</head>
<body>
    <h1>Lista de Productos</h1>
    <a href="crear_producto.php">Añadir Producto</a>
    <table border="1">
        <tr>
            <th>ID</th>
            <th>Nombre</th>
            <th>Descripción</th>
            <th>Precio</th>
            <th>Stock</th>
            <th>Acciones</th>
        </tr>
        <?php while ($row = $result->fetch_assoc()): ?>
        <tr>
            <td><?= $row['id'] ?></td>
            <td><?= $row['nombre'] ?></td>
            <td><?= $row['descripcion'] ?></td>
            <td><?= $row['precio'] ?></td>
            <td><?= $row['stock'] ?></td>
            <td>
                <a href="editar_producto.php?id=<?= $row['id'] ?>">Editar</a>
                <a href="eliminar_producto.php?id=<?= $row['id'] ?>" onclick="return confirm('¿Seguro que deseas eliminar este producto?');">Eliminar</a>
            </td>
        </tr>
        <?php endwhile; ?>
    </table>
</body>
</html>
