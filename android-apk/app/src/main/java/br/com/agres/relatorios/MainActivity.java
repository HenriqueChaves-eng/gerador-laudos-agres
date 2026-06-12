package br.com.agres.relatorios;

import android.Manifest;
import android.app.Activity;
import android.content.ClipData;
import android.content.ContentValues;
import android.content.Intent;
import android.content.pm.ActivityInfo;
import android.content.pm.PackageManager;
import android.content.pm.ResolveInfo;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.util.Base64;
import android.webkit.GeolocationPermissions;
import android.webkit.JavascriptInterface;
import android.webkit.MimeTypeMap;
import android.webkit.PermissionRequest;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebChromeClient.FileChooserParams;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class MainActivity extends Activity {
    private static final String APP_ORIGIN = "https://agres-offline.local/";
    private static final int FILE_CHOOSER_REQUEST = 501;
    private static final int PERMISSIONS_REQUEST = 502;
    private static final int MAX_JSON_CHARS = 100 * 1024 * 1024;
    private static final int MAX_IMAGE_BASE64_CHARS = 36 * 1024 * 1024;

    private WebView webView;
    private ValueCallback<Uri[]> filePathCallback;
    private Uri cameraPhotoUri;
    private File cameraPhotoFile;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        cleanupCacheFiles();
        requestRuntimePermissions();

        webView = new WebView(this);
        setContentView(webView);
        configureWebView();
        webView.loadUrl(APP_ORIGIN + "index.html");
    }

    private void requestRuntimePermissions() {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.M) return;

        ArrayList<String> permissions = new ArrayList<>();
        addPermissionIfMissing(permissions, Manifest.permission.CAMERA);
        addPermissionIfMissing(permissions, Manifest.permission.RECORD_AUDIO);
        addPermissionIfMissing(permissions, Manifest.permission.ACCESS_FINE_LOCATION);
        addPermissionIfMissing(permissions, Manifest.permission.ACCESS_COARSE_LOCATION);
        if (Build.VERSION.SDK_INT <= Build.VERSION_CODES.P) {
            addPermissionIfMissing(permissions, Manifest.permission.WRITE_EXTERNAL_STORAGE);
        }

        if (!permissions.isEmpty()) {
            requestPermissions(permissions.toArray(new String[0]), PERMISSIONS_REQUEST);
        }
    }

    private void addPermissionIfMissing(ArrayList<String> permissions, String permission) {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.M && checkSelfPermission(permission) != PackageManager.PERMISSION_GRANTED) {
            permissions.add(permission);
        }
    }

    private void configureWebView() {
        WebView.setWebContentsDebuggingEnabled(false);

        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setGeolocationEnabled(true);
        settings.setMediaPlaybackRequiresUserGesture(false);
        settings.setAllowFileAccess(false);
        settings.setAllowContentAccess(true);
        settings.setJavaScriptCanOpenWindowsAutomatically(true);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
            settings.setSafeBrowsingEnabled(true);
        }
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.LOLLIPOP) {
            settings.setMixedContentMode(WebSettings.MIXED_CONTENT_NEVER_ALLOW);
        }

        webView.addJavascriptInterface(new AndroidBridge(), "AgresAndroid");
        webView.setWebViewClient(new OfflineWebViewClient());
        webView.setWebChromeClient(new OfflineChromeClient());
    }

    private class OfflineWebViewClient extends WebViewClient {
        @Override
        public boolean shouldOverrideUrlLoading(WebView view, WebResourceRequest request) {
            return openExternalIfNeeded(request.getUrl());
        }

        @Override
        public boolean shouldOverrideUrlLoading(WebView view, String url) {
            return openExternalIfNeeded(Uri.parse(url));
        }

        @Override
        public WebResourceResponse shouldInterceptRequest(WebView view, WebResourceRequest request) {
            return assetResponse(request.getUrl());
        }

        @Override
        public WebResourceResponse shouldInterceptRequest(WebView view, String url) {
            return assetResponse(Uri.parse(url));
        }
    }

    private boolean openExternalIfNeeded(Uri uri) {
        if (uri == null || "agres-offline.local".equals(uri.getHost())) {
            return false;
        }
        try {
            Intent intent = new Intent(Intent.ACTION_VIEW, uri);
            startActivity(intent);
            return true;
        } catch (Exception error) {
            return true;
        }
    }

    private WebResourceResponse assetResponse(Uri uri) {
        if (uri == null || !"agres-offline.local".equals(uri.getHost())) {
            return null;
        }

        String path = uri.getPath();
        if (path == null || "/".equals(path) || path.trim().isEmpty()) {
            path = "/index.html";
        }
        path = path.replaceFirst("^/+", "");
        if (path.contains("..")) {
            return null;
        }

        try {
            InputStream stream = getAssets().open("offline/" + path);
            return new WebResourceResponse(mimeType(path), "UTF-8", stream);
        } catch (Exception error) {
            return null;
        }
    }

    private String mimeType(String path) {
        String lower = path.toLowerCase(Locale.ROOT);
        if (lower.endsWith(".html")) return "text/html";
        if (lower.endsWith(".js")) return "application/javascript";
        if (lower.endsWith(".json") || lower.endsWith(".webmanifest")) return "application/manifest+json";
        if (lower.endsWith(".png")) return "image/png";
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) return "image/jpeg";
        if (lower.endsWith(".css")) return "text/css";
        return "application/octet-stream";
    }

    private class OfflineChromeClient extends WebChromeClient {
        @Override
        public void onPermissionRequest(final PermissionRequest request) {
            runOnUiThread(() -> {
                Uri origin = request.getOrigin();
                if (origin == null || !"agres-offline.local".equals(origin.getHost())) {
                    request.deny();
                    return;
                }
                ArrayList<String> allowed = new ArrayList<>();
                for (String resource : request.getResources()) {
                    if (PermissionRequest.RESOURCE_AUDIO_CAPTURE.equals(resource)
                        || PermissionRequest.RESOURCE_VIDEO_CAPTURE.equals(resource)) {
                        allowed.add(resource);
                    }
                }
                if (allowed.isEmpty()) {
                    request.deny();
                } else {
                    request.grant(allowed.toArray(new String[0]));
                }
            });
        }

        @Override
        public void onGeolocationPermissionsShowPrompt(String origin, GeolocationPermissions.Callback callback) {
            Uri uri = Uri.parse(origin);
            boolean allowed = uri != null && "agres-offline.local".equals(uri.getHost());
            callback.invoke(origin, allowed, false);
        }

        @Override
        public boolean onShowFileChooser(WebView webView, ValueCallback<Uri[]> callback, FileChooserParams params) {
            if (filePathCallback != null) {
                filePathCallback.onReceiveValue(null);
            }
            filePathCallback = callback;

            Intent contentIntent = new Intent(Intent.ACTION_GET_CONTENT);
            contentIntent.addCategory(Intent.CATEGORY_OPENABLE);
            boolean imageChooser = acceptsImage(params);
            contentIntent.setType(acceptedMimeType(params));
            contentIntent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, params.getMode() == FileChooserParams.MODE_OPEN_MULTIPLE);

            ArrayList<Intent> initialIntents = new ArrayList<>();
            Intent cameraIntent = params.isCaptureEnabled() && imageChooser ? createCameraIntent() : null;
            if (cameraIntent != null) {
                initialIntents.add(cameraIntent);
            }

            Intent chooser;
            if (params.isCaptureEnabled() && cameraIntent != null) {
                chooser = cameraIntent;
            } else {
                chooser = Intent.createChooser(contentIntent, "Selecionar arquivo");
                if (!initialIntents.isEmpty()) {
                    chooser.putExtra(Intent.EXTRA_INITIAL_INTENTS, initialIntents.toArray(new Intent[0]));
                }
            }

            try {
                startActivityForResult(chooser, FILE_CHOOSER_REQUEST);
            } catch (Exception error) {
                filePathCallback.onReceiveValue(null);
                filePathCallback = null;
                return false;
            }
            return true;
        }
    }

    private boolean acceptsImage(WebChromeClient.FileChooserParams params) {
        String[] acceptTypes = params.getAcceptTypes();
        if (acceptTypes == null || acceptTypes.length == 0) {
            return false;
        }
        for (String acceptType : acceptTypes) {
            if (acceptType != null && acceptType.trim().startsWith("image/")) {
                return true;
            }
        }
        return false;
    }

    private String acceptedMimeType(WebChromeClient.FileChooserParams params) {
        String[] acceptTypes = params.getAcceptTypes();
        if (acceptTypes != null) {
            for (String acceptType : acceptTypes) {
                if (acceptType != null && acceptType.trim().startsWith("image/")) {
                    return "image/*";
                }
                if (acceptType != null && !acceptType.trim().isEmpty()) {
                    return acceptType;
                }
            }
        }
        return "*/*";
    }

    private Intent createCameraIntent() {
        Intent intent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
        if (intent.resolveActivity(getPackageManager()) == null) {
            return null;
        }

        try {
            cleanupCacheFiles();
            File cameraDir = new File(getCacheDir(), "camera");
            if (!cameraDir.exists()) cameraDir.mkdirs();
            String stamp = new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.ROOT).format(new Date());
            cameraPhotoFile = new File(cameraDir, "agres_" + stamp + ".jpg");
            cameraPhotoUri = Uri.parse("content://" + getPackageName() + ".camera/" + cameraPhotoFile.getName());
            intent.putExtra(MediaStore.EXTRA_OUTPUT, cameraPhotoUri);
            intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION);
            intent.setClipData(ClipData.newUri(getContentResolver(), "Agres Foto", cameraPhotoUri));
            return intent;
        } catch (Exception error) {
            cameraPhotoFile = null;
            cameraPhotoUri = null;
            return null;
        }
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);
        if (requestCode != FILE_CHOOSER_REQUEST || filePathCallback == null) {
            return;
        }

        ArrayList<Uri> uris = new ArrayList<>();
        if (resultCode == RESULT_OK) {
            if (data != null && data.getClipData() != null) {
                ClipData clipData = data.getClipData();
                for (int index = 0; index < clipData.getItemCount(); index += 1) {
                    uris.add(clipData.getItemAt(index).getUri());
                }
            } else if (data != null && data.getData() != null) {
                uris.add(data.getData());
            } else if (cameraPhotoUri != null && cameraPhotoFile != null && cameraPhotoFile.exists() && cameraPhotoFile.length() > 0) {
                uris.add(cameraPhotoUri);
            }
        }

        filePathCallback.onReceiveValue(uris.isEmpty() ? null : uris.toArray(new Uri[0]));
        filePathCallback = null;
        cameraPhotoUri = null;
        cameraPhotoFile = null;
    }

    @Override
    public void onBackPressed() {
        if (webView != null && webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }

    @Override
    protected void onDestroy() {
        if (webView != null) {
            webView.stopLoading();
            webView.removeJavascriptInterface("AgresAndroid");
            webView.setWebChromeClient(null);
            webView.setWebViewClient(null);
            webView.destroy();
            webView = null;
        }
        super.onDestroy();
    }

    public class AndroidBridge {
        @JavascriptInterface
        public boolean saveImage(String dataUrl, String requestedName) {
            try {
                DecodedImage image = decodeImage(dataUrl, requestedName);
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                    return saveImageModern(image);
                }
                return saveImageLegacy(image);
            } catch (Exception error) {
                return false;
            }
        }

        @JavascriptInterface
        public boolean saveJson(String jsonText, String requestedName) {
            try {
                if (jsonText == null || jsonText.length() == 0 || jsonText.length() > MAX_JSON_CHARS) return false;
                String fileName = normalizedJsonName(requestedName);
                byte[] bytes = jsonText.getBytes(StandardCharsets.UTF_8);
                if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                    return saveJsonModern(bytes, fileName);
                }
                return saveJsonLegacy(bytes, fileName);
            } catch (Exception error) {
                return false;
            }
        }

        @JavascriptInterface
        public boolean shareJson(String jsonText, String requestedName) {
            try {
                if (jsonText == null || jsonText.length() == 0 || jsonText.length() > MAX_JSON_CHARS) return false;
                String fileName = normalizedJsonName(requestedName);
                byte[] bytes = jsonText.getBytes(StandardCharsets.UTF_8);
                Uri uri = writeJsonShareFile(bytes, fileName);

                Intent shareIntent = new Intent(Intent.ACTION_SEND);
                shareIntent.setType("application/octet-stream");
                shareIntent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{"application/json", "application/octet-stream"});
                shareIntent.putExtra(Intent.EXTRA_STREAM, uri);
                shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                shareIntent.setClipData(ClipData.newUri(getContentResolver(), "Pacote Relatorio Offline", uri));

                List<ResolveInfo> shareTargets = getPackageManager().queryIntentActivities(
                    shareIntent,
                    PackageManager.MATCH_DEFAULT_ONLY
                );
                for (ResolveInfo target : shareTargets) {
                    if (target.activityInfo != null && target.activityInfo.packageName != null) {
                        grantUriPermission(
                            target.activityInfo.packageName,
                            uri,
                            Intent.FLAG_GRANT_READ_URI_PERMISSION
                        );
                    }
                }

                Intent chooser = Intent.createChooser(shareIntent, "Compartilhar pacote relatorio offline");
                chooser.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                runOnUiThread(() -> {
                    try {
                        startActivity(chooser);
                    } catch (Exception ignored) {
                    }
                });
                return true;
            } catch (Exception error) {
                return false;
            }
        }

        @JavascriptInterface
        public void setSignatureLandscape(boolean enabled) {
            runOnUiThread(() -> setRequestedOrientation(
                enabled
                    ? ActivityInfo.SCREEN_ORIENTATION_LANDSCAPE
                    : ActivityInfo.SCREEN_ORIENTATION_UNSPECIFIED
            ));
        }
    }

    private static class DecodedImage {
        byte[] bytes;
        String mimeType;
        String fileName;
    }

    private DecodedImage decodeImage(String dataUrl, String requestedName) {
        String raw = dataUrl == null ? "" : dataUrl.trim();
        int comma = raw.indexOf(',');
        if (!raw.startsWith("data:") || comma < 0) {
            throw new IllegalArgumentException("Imagem inválida.");
        }

        String header = raw.substring(5, comma);
        String mime = header.split(";")[0];
        if (mime == null || !mime.startsWith("image/")) {
            mime = "image/jpeg";
        }

        String base64 = raw.substring(comma + 1);
        if (base64.length() > MAX_IMAGE_BASE64_CHARS) {
            throw new IllegalArgumentException("Imagem acima do limite permitido.");
        }
        DecodedImage image = new DecodedImage();
        image.bytes = Base64.decode(base64, Base64.DEFAULT);
        image.mimeType = mime;
        image.fileName = normalizedImageName(requestedName, mime);
        return image;
    }

    private String normalizedImageName(String requestedName, String mimeType) {
        String name = requestedName == null ? "" : requestedName.trim();
        name = name.replaceAll("[\\\\/:*?\"<>|]+", "_").replaceAll("\\s+", "_");
        if (name.isEmpty()) {
            name = "foto_agres_" + new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.ROOT).format(new Date());
        }
        String extension = MimeTypeMap.getSingleton().getExtensionFromMimeType(mimeType);
        if (extension == null || extension.trim().isEmpty()) {
            extension = "jpg";
        }
        if (!name.toLowerCase(Locale.ROOT).endsWith("." + extension.toLowerCase(Locale.ROOT))) {
            name = name.replaceFirst("\\.[^.]+$", "") + "." + extension;
        }
        return name;
    }

    private boolean saveImageModern(DecodedImage image) throws Exception {
        ContentValues values = new ContentValues();
        values.put(MediaStore.Images.Media.DISPLAY_NAME, image.fileName);
        values.put(MediaStore.Images.Media.MIME_TYPE, image.mimeType);
        values.put(MediaStore.Images.Media.RELATIVE_PATH, Environment.DIRECTORY_PICTURES + "/Agres Relatorios");
        values.put(MediaStore.Images.Media.IS_PENDING, 1);

        Uri uri = getContentResolver().insert(MediaStore.Images.Media.EXTERNAL_CONTENT_URI, values);
        if (uri == null) return false;

        boolean completed = false;
        try {
            try (OutputStream output = getContentResolver().openOutputStream(uri)) {
                if (output == null) return false;
                output.write(image.bytes);
            }
            values.clear();
            values.put(MediaStore.Images.Media.IS_PENDING, 0);
            getContentResolver().update(uri, values, null, null);
            completed = true;
            return true;
        } finally {
            if (!completed) getContentResolver().delete(uri, null, null);
        }
    }

    private boolean saveImageLegacy(DecodedImage image) throws Exception {
        File dir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_PICTURES), "Agres Relatorios");
        if (!dir.exists() && !dir.mkdirs()) return false;
        File file = new File(dir, image.fileName);
        try (FileOutputStream output = new FileOutputStream(file)) {
            output.write(image.bytes);
        }
        sendBroadcast(new Intent(Intent.ACTION_MEDIA_SCANNER_SCAN_FILE, Uri.fromFile(file)));
        return true;
    }

    private String normalizedJsonName(String requestedName) {
        String name = requestedName == null ? "" : requestedName.trim();
        name = name.replaceAll("[\\\\/:*?\"<>|]+", "_").replaceAll("\\s+", "_");
        if (name.isEmpty()) {
            name = "pacote_relatorio_offline_" + new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.ROOT).format(new Date());
        }
        if (!name.toLowerCase(Locale.ROOT).endsWith(".json")) {
            name = name.replaceFirst("\\.[^.]+$", "") + ".json";
        }
        return name;
    }

    private boolean saveJsonModern(byte[] bytes, String fileName) throws Exception {
        ContentValues values = new ContentValues();
        values.put(MediaStore.MediaColumns.DISPLAY_NAME, fileName);
        values.put(MediaStore.MediaColumns.MIME_TYPE, "application/json");
        values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/Agres Relatorios");
        values.put(MediaStore.MediaColumns.IS_PENDING, 1);

        Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
        if (uri == null) return false;

        boolean completed = false;
        try {
            try (OutputStream output = getContentResolver().openOutputStream(uri)) {
                if (output == null) return false;
                output.write(bytes);
            }
            values.clear();
            values.put(MediaStore.MediaColumns.IS_PENDING, 0);
            getContentResolver().update(uri, values, null, null);
            completed = true;
            return true;
        } finally {
            if (!completed) getContentResolver().delete(uri, null, null);
        }
    }

    private boolean saveJsonLegacy(byte[] bytes, String fileName) throws Exception {
        File dir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "Agres Relatorios");
        if (!dir.exists() && !dir.mkdirs()) return false;
        File file = new File(dir, fileName);
        try (FileOutputStream output = new FileOutputStream(file)) {
            output.write(bytes);
        }
        sendBroadcast(new Intent(Intent.ACTION_MEDIA_SCANNER_SCAN_FILE, Uri.fromFile(file)));
        return true;
    }

    private Uri writeJsonShareFile(byte[] bytes, String fileName) throws Exception {
        cleanupCacheFiles();
        File dir = new File(getCacheDir(), "camera");
        if (!dir.exists() && !dir.mkdirs()) {
            throw new IllegalStateException("Não foi possível preparar o compartilhamento.");
        }
        File file = new File(dir, fileName);
        try (FileOutputStream output = new FileOutputStream(file, false)) {
            output.write(bytes);
        }
        return new Uri.Builder()
            .scheme("content")
            .authority(getPackageName() + ".camera")
            .appendPath(fileName)
            .build();
    }

    private void cleanupCacheFiles() {
        File dir = new File(getCacheDir(), "camera");
        File[] files = dir.listFiles();
        if (files == null) return;
        long cutoff = System.currentTimeMillis() - (24L * 60L * 60L * 1000L);
        for (File file : files) {
            if (file.isFile() && file.lastModified() < cutoff) {
                file.delete();
            }
        }
    }
}
