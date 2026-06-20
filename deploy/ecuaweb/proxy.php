<?php
/**
 * Proxy PHP - fallback si mod_proxy no esta disponible en Ecuaweb.
 *
 * Subir a:   audit-ia.ec/sri_robot_audit/proxy.php
 * Y reemplazar el .htaccess por:
 *
 *   RewriteEngine On
 *   RewriteCond %{REQUEST_FILENAME} !-f
 *   RewriteCond %{REQUEST_FILENAME} !-d
 *   RewriteRule ^(.*)$ proxy.php?p=$1 [QSA,L]
 *
 * Resultado: cada peticion a /sri_robot_audit/<path> entra a este script,
 * que llama al backend de Render con cURL y reenvia la respuesta tal cual.
 *
 * Soporta: GET, POST, PUT, DELETE, cookies, redirect interno (Location),
 * binary content (e.g. el .exe).
 */

declare(strict_types=1);

const BACKEND = 'https://sri-robot-audit-ik01.onrender.com';

// Path solicitado al proxy (e.g. "admin/login").
$path = isset($_GET['p']) ? (string) $_GET['p'] : '';

// Si /landing -> ir a / en el backend (la landing vive en raiz en Render).
if ($path === 'landing' || $path === 'landing/') {
    $path = '';
}

// Query string original sin el parametro 'p' (que es interno del proxy).
parse_str($_SERVER['QUERY_STRING'] ?? '', $qs);
unset($qs['p']);
$qsString = http_build_query($qs);

$targetUrl = rtrim(BACKEND, '/') . '/' . ltrim($path, '/');
if ($qsString !== '') {
    $targetUrl .= '?' . $qsString;
}

$ch = curl_init($targetUrl);
curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
curl_setopt($ch, CURLOPT_FOLLOWLOCATION, false);  // Reenviamos los 3xx tal cual
curl_setopt($ch, CURLOPT_HEADER, true);
curl_setopt($ch, CURLOPT_BINARYTRANSFER, true);
curl_setopt($ch, CURLOPT_TIMEOUT, 120);
curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, 10);
curl_setopt($ch, CURLOPT_CUSTOMREQUEST, $_SERVER['REQUEST_METHOD']);

// Cuerpo de la request (POST, PUT) — leer raw input.
$body = file_get_contents('php://input');
if ($body !== '' && $body !== false) {
    curl_setopt($ch, CURLOPT_POSTFIELDS, $body);
}

// Forward de headers utiles.
$headers = [];
foreach ($_SERVER as $key => $value) {
    if (strpos($key, 'HTTP_') === 0) {
        $headerName = str_replace('_', '-', substr($key, 5));
        // No reenviar Host (lo setea curl) ni cookies (las pasamos abajo).
        if (in_array(strtolower($headerName), ['host', 'cookie', 'connection'], true)) {
            continue;
        }
        $headers[] = $headerName . ': ' . $value;
    }
}
if (isset($_SERVER['CONTENT_TYPE'])) {
    $headers[] = 'Content-Type: ' . $_SERVER['CONTENT_TYPE'];
}
$headers[] = 'X-Forwarded-Proto: https';
$headers[] = 'X-Forwarded-Host: ' . ($_SERVER['HTTP_HOST'] ?? 'audit-ia.ec');
curl_setopt($ch, CURLOPT_HTTPHEADER, $headers);

// Forward de cookies del cliente al backend.
if (!empty($_COOKIE)) {
    $cookieStr = '';
    foreach ($_COOKIE as $k => $v) {
        $cookieStr .= rawurlencode((string) $k) . '=' . rawurlencode((string) $v) . '; ';
    }
    curl_setopt($ch, CURLOPT_COOKIE, rtrim($cookieStr, '; '));
}

$response = curl_exec($ch);
if ($response === false) {
    http_response_code(502);
    header('Content-Type: text/plain; charset=utf-8');
    echo 'Bad gateway: ' . curl_error($ch);
    curl_close($ch);
    exit;
}

$statusCode = curl_getinfo($ch, CURLINFO_RESPONSE_CODE);
$headerSize = curl_getinfo($ch, CURLINFO_HEADER_SIZE);
$rawHeaders = substr($response, 0, $headerSize);
$responseBody = substr($response, $headerSize);
curl_close($ch);

http_response_code($statusCode);

// Reenviar headers seleccionados del backend al cliente.
$forwardable = [
    'content-type',
    'content-disposition',
    'content-length',
    'location',
    'set-cookie',
    'cache-control',
    'etag',
    'last-modified',
    'x-frame-options',
    'x-content-type-options',
];
foreach (explode("\r\n", $rawHeaders) as $headerLine) {
    $colonPos = strpos($headerLine, ':');
    if ($colonPos === false) {
        continue;
    }
    $name = strtolower(trim(substr($headerLine, 0, $colonPos)));
    if (in_array($name, $forwardable, true)) {
        // header() reemplaza por defecto, salvo Set-Cookie (multiple).
        $replace = ($name !== 'set-cookie');
        header($headerLine, $replace);
    }
}

echo $responseBody;
