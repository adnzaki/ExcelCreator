<?php

require_once 'vendor/autoload.php';

header('Content-Type: application/json');

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['error' => 'Only POST method is allowed']);
    exit;
}

if (!isset($_FILES['excel'])) {
    http_response_code(400);
    echo json_encode(['error' => 'No file uploaded']);
    exit;
}

$file = $_FILES['excel'];

if ($file['error'] !== UPLOAD_ERR_OK) {
    http_response_code(500);
    echo json_encode(['error' => 'Upload failed']);
    exit;
}

// Save to folder uploads/
$targetPath = __DIR__ . '/uploads/' . basename($file['name']);
move_uploaded_file($file['tmp_name'], $targetPath);

// Reads Excel file
$reader = new \ExcelTools\Reader();
$reader->loadFromFile($targetPath);
$data = $reader->getSheetData(true); // true jika header ingin jadi key

// Cleanup files after upload
unlink($targetPath);

// Send response as JSON
echo json_encode($data);