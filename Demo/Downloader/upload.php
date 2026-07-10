<?php
/**
 * czSocket Upload Endpoint
 * 
 * Place this file in your Laragon www folder:
 *   C:\laragon\www\upload.php
 * 
 * Access at: http://localhost/upload.php
 * 
 * Accepts multipart/form-data file uploads.
 * Saves to ./uploads/ subfolder.
 * Returns JSON response.
 */

header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');

// Create uploads directory if it doesn't exist
$uploadDir = __DIR__ . '/uploads/';
if (!is_dir($uploadDir)) {
    mkdir($uploadDir, 0777, true);
}

// Handle POST upload
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (empty($_FILES)) {
        http_response_code(400);
        echo json_encode([
            'success' => false,
            'error' => 'No file uploaded'
        ]);
        exit;
    }

    $results = [];
    foreach ($_FILES as $fieldName => $file) {
        if ($file['error'] === UPLOAD_ERR_OK) {
            $safeName = preg_replace('/[^a-zA-Z0-9._-]/', '_', basename($file['name']));
            // Avoid overwrite: add timestamp if file exists
            if (file_exists($uploadDir . $safeName)) {
                $ext = pathinfo($safeName, PATHINFO_EXTENSION);
                $base = pathinfo($safeName, PATHINFO_FILENAME);
                $safeName = $base . '_' . time() . '.' . $ext;
            }
            $destPath = $uploadDir . $safeName;
            if (move_uploaded_file($file['tmp_name'], $destPath)) {
                $results[] = [
                    'field' => $fieldName,
                    'name' => $safeName,
                    'original_name' => $file['name'],
                    'size' => $file['size'],
                    'size_human' => formatBytes($file['size']),
                    'type' => $file['type'],
                    'path' => '/uploads/' . $safeName
                ];
            } else {
                $results[] = [
                    'field' => $fieldName,
                    'name' => $file['name'],
                    'error' => 'Failed to move uploaded file'
                ];
            }
        } else {
            $results[] = [
                'field' => $fieldName,
                'name' => $file['name'],
                'error' => uploadErrorMessage($file['error'])
            ];
        }
    }

    echo json_encode([
        'success' => true,
        'files' => $results,
        'server' => 'Laragon/PHP ' . phpversion(),
        'timestamp' => date('Y-m-d H:i:s')
    ], JSON_PRETTY_PRINT);
    exit;
}

// Handle GET - show upload form / status
if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    // List existing uploads
    $files = [];
    if (is_dir($uploadDir)) {
        foreach (scandir($uploadDir) as $f) {
            if ($f !== '.' && $f !== '..') {
                $files[] = [
                    'name' => $f,
                    'size' => filesize($uploadDir . $f),
                    'size_human' => formatBytes(filesize($uploadDir . $f)),
                    'modified' => date('Y-m-d H:i:s', filemtime($uploadDir . $f))
                ];
            }
        }
    }
    echo json_encode([
        'status' => 'ready',
        'endpoint' => 'POST /upload.php',
        'field_name' => 'file',
        'max_upload_size' => ini_get('upload_max_filesize'),
        'uploaded_files' => $files
    ], JSON_PRETTY_PRINT);
    exit;
}

function formatBytes($bytes) {
    if ($bytes < 1024) return $bytes . ' B';
    if ($bytes < 1048576) return round($bytes / 1024, 1) . ' KB';
    if ($bytes < 1073741824) return round($bytes / 1048576, 1) . ' MB';
    return round($bytes / 1073741824, 1) . ' GB';
}

function uploadErrorMessage($code) {
    $errors = [
        UPLOAD_ERR_INI_SIZE => 'File exceeds upload_max_filesize',
        UPLOAD_ERR_FORM_SIZE => 'File exceeds MAX_FILE_SIZE',
        UPLOAD_ERR_PARTIAL => 'File was only partially uploaded',
        UPLOAD_ERR_NO_FILE => 'No file was uploaded',
        UPLOAD_ERR_NO_TMP_DIR => 'Missing temp folder',
        UPLOAD_ERR_CANT_WRITE => 'Failed to write to disk',
        UPLOAD_ERR_EXTENSION => 'Upload stopped by extension'
    ];
    return $errors[$code] ?? 'Unknown error';
}
