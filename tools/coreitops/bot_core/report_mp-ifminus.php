<?php
// Proteksi HTTP Basic Auth
$valid_user = "admin-era";
$valid_pass = "HUG8wQdfa6DWhY4n";

if (!isset($_SERVER['PHP_AUTH_USER']) || 
    $_SERVER['PHP_AUTH_USER'] !== $valid_user || 
    $_SERVER['PHP_AUTH_PW'] !== $valid_pass) {
    
    header('WWW-Authenticate: Basic realm="Restricted Area"');
    header('HTTP/1.0 401 Unauthorized');
    echo json_encode(["status" => "error", "message" => "Akses Ditolak: Kredensial Salah"]);
    exit;
}

header('Content-Type: application/json');

try {
    // Koneksi lokal ke database via PDO
    $host = '172.20.82.44';
    $db   = 'apis_2022';
    $user = 'eraclub';
    $pass = 'mariadb2019@eraclub.com';
    $charset = 'utf8mb4';

    $dsn = "mysql:host=$host;dbname=$db;charset=$charset"; // Ganti 'mysql' ke 'pgsql' jika PostgreSQL
    $pdo = new PDO($dsn, $user, $pass, [
        PDO::ATTR_ERRMODE => PDO_ERRMODE_EXCEPTION,
        PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC
    ]);

    // Baca teks query yang dikirim oleh Bot via POST
    $input = json_decode(file_get_contents('php://input'), true);
    $sql_query = $input['query'] ?? '';

    if (empty($sql_query)) {
        echo json_encode(["status" => "error", "message" => "Query kosong"]);
        exit;
    }

    // Eksekusi query
    $stmt = $pdo->query($sql_query);
    $data = $stmt->fetchAll();

    // Kembalikan hasil query ke Bot
    echo json_encode([
        "status" => "success",
        "total_rows" => count($data),
        "data" => $data
    ]);

} catch (Exception $e) {
    http_response_code(500);
    echo json_encode(["status" => "error", "message" => $e->getMessage()]);
}
?>