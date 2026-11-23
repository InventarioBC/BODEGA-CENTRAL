package cl.grupogyl.bodegacentral;

import android.Manifest;
import android.annotation.SuppressLint;
import android.app.Activity;
import android.content.Intent;
import android.content.pm.PackageManager;
import android.graphics.Bitmap;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.view.View;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceError;
import android.webkit.WebResourceRequest;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.webkit.WebViewClient;
import android.widget.ProgressBar;
import android.widget.Toast;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.app.ActivityCompat;
import androidx.core.content.ContextCompat;
import androidx.core.content.FileProvider;
import androidx.swiperefreshlayout.widget.SwipeRefreshLayout;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

public class MainActivity extends AppCompatActivity {

    private static final String PRIMARY_URL = "http://192.168.1.18:8080/?v=31";
    private static final String FALLBACK_URL = "https://florists-spin-guards-phillips.trycloudflare.com/?v=31";
    private static final int PERMISSION_REQ_CODE = 1001;
    private static final int FILE_CHOOSER_REQ_CODE = 1002;

    private WebView webView;
    private SwipeRefreshLayout swipeRefreshLayout;
    private ProgressBar progressBar;

    private ValueCallback<Uri[]> fileUploadCallback;
    private String currentPhotoPath;

    @SuppressLint("SetJavaScriptEnabled")
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        webView = findViewById(R.id.webView);
        swipeRefreshLayout = findViewById(R.id.swipeRefreshLayout);
        progressBar = findViewById(R.id.progressBar);

        setupWebView();
        setupSwipeRefresh();
        checkAndRequestPermissions();

        // Cargar URL principal
        webView.loadUrl(PRIMARY_URL);
    }

    @SuppressLint("SetJavaScriptEnabled")
    private void setupWebView() {
        WebSettings settings = webView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setAllowFileAccess(true);
        settings.setAllowContentAccess(true);
        settings.setLoadWithOverviewMode(true);
        settings.setUseWideViewPort(true);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);
        settings.setMediaPlaybackRequiresUserGesture(false);

        // Habilitar caché local agresivo para funcionamiento rápido
        settings.setCacheMode(WebSettings.LOAD_DEFAULT);

        webView.setWebViewClient(new WebViewClient() {
            @Override
            public void onPageStarted(WebView view, String url, Bitmap favicon) {
                progressBar.setVisibility(View.VISIBLE);
                super.onPageStarted(view, url, favicon);
            }

            @Override
            public void onPageFinished(WebView view, String url) {
                progressBar.setVisibility(View.GONE);
                swipeRefreshLayout.setRefreshing(false);
                super.onPageFinished(view, url);
            }

            @Override
            public void onReceivedError(WebView view, WebResourceRequest request, WebResourceError error) {
                if (request.isForMainFrame()) {
                    String failedUrl = request.getUrl().toString();
                    // Si falló el enlace local de Wi-Fi, intentar automáticamente el túnel Cloudflare
                    if (failedUrl.contains("192.168.1.18")) {
                        view.loadUrl(FALLBACK_URL);
                    } else {
                        showErrorHtml(view);
                    }
                }
                super.onReceivedError(view, request, error);
            }
        });

        webView.setWebChromeClient(new WebChromeClient() {
            @Override
            public void onProgressChanged(WebView view, int newProgress) {
                progressBar.setProgress(newProgress);
                if (newProgress == 100) {
                    progressBar.setVisibility(View.GONE);
                }
            }

            // Manejador nativo para tomar fotos con la cámara o seleccionar de galería
            @Override
            public boolean onShowFileChooser(WebView webView, ValueCallback<Uri[]> filePathCallback, FileChooserParams fileChooserParams) {
                if (fileUploadCallback != null) {
                    fileUploadCallback.onReceiveValue(null);
                }
                fileUploadCallback = filePathCallback;

                Intent takePictureIntent = new Intent(MediaStore.ACTION_IMAGE_CAPTURE);
                File photoFile = null;
                try {
                    photoFile = createImageFile();
                } catch (IOException ex) {
                    // Ignorar
                }

                if (photoFile != null) {
                    Uri photoURI = FileProvider.getUriForFile(MainActivity.this,
                            getApplicationContext().getPackageName() + ".provider",
                            photoFile);
                    takePictureIntent.putExtra(MediaStore.EXTRA_OUTPUT, photoURI);
                }

                Intent contentSelectionIntent = new Intent(Intent.ACTION_GET_CONTENT);
                contentSelectionIntent.addCategory(Intent.CATEGORY_OPENABLE);
                contentSelectionIntent.setType("image/*");

                Intent[] intentArray;
                if (takePictureIntent.resolveActivity(getPackageManager()) != null) {
                    intentArray = new Intent[]{takePictureIntent};
                } else {
                    intentArray = new Intent[0];
                }

                Intent chooserIntent = new Intent(Intent.ACTION_CHOOSER);
                chooserIntent.putExtra(Intent.EXTRA_INTENT, contentSelectionIntent);
                chooserIntent.putExtra(Intent.EXTRA_TITLE, "Tomar Foto o Seleccionar Imagen");
                chooserIntent.putExtra(Intent.EXTRA_INITIAL_INTENTS, intentArray);

                startActivityForResult(chooserIntent, FILE_CHOOSER_REQ_CODE);
                return true;
            }
        });
    }

    private void setupSwipeRefresh() {
        swipeRefreshLayout.setColorSchemeResources(R.color.primary, R.color.accent);
        swipeRefreshLayout.setOnRefreshListener(() -> {
            webView.reload();
        });
    }

    private void checkAndRequestPermissions() {
        if (ContextCompat.checkSelfPermission(this, Manifest.permission.CAMERA) != PackageManager.PERMISSION_GRANTED) {
            ActivityCompat.requestPermissions(this,
                    new String[]{Manifest.permission.CAMERA},
                    PERMISSION_REQ_CODE);
        }
    }

    private File createImageFile() throws IOException {
        String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(new Date());
        String imageFileName = "BODEGA_" + timeStamp + "_";
        File storageDir = getExternalFilesDir(Environment.DIRECTORY_PICTURES);
        File image = File.createTempFile(imageFileName, ".jpg", storageDir);
        currentPhotoPath = image.getAbsolutePath();
        return image;
    }

    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        if (requestCode == FILE_CHOOSER_REQ_CODE) {
            if (fileUploadCallback == null) {
                super.onActivityResult(requestCode, resultCode, data);
                return;
            }

            Uri[] results = null;
            if (resultCode == Activity.RESULT_OK) {
                if (data == null || data.getData() == null) {
                    // Foto tomada por cámara
                    if (currentPhotoPath != null) {
                        results = new Uri[]{Uri.fromFile(new File(currentPhotoPath))};
                    }
                } else {
                    String dataString = data.getDataString();
                    if (dataString != null) {
                        results = new Uri[]{Uri.parse(dataString)};
                    }
                }
            }

            fileUploadCallback.onReceiveValue(results);
            fileUploadCallback = null;
            return;
        }
        super.onActivityResult(requestCode, resultCode, data);
    }

    @Override
    public void onBackPressed() {
        if (webView.canGoBack()) {
            webView.goBack();
        } else {
            super.onBackPressed();
        }
    }

    private void showErrorHtml(WebView view) {
        String errorHtml = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
                "<meta name='viewport' content='width=device-width, initial-scale=1.0'>" +
                "<style>" +
                "body { font-family: -apple-system, sans-serif; background: #EAEFF4; text-align: center; padding: 40px 20px; color: #2D3748; }" +
                ".card { background: white; border-radius: 20px; padding: 30px 20px; box-shadow: 0 10px 25px rgba(0,0,0,0.08); }" +
                "h2 { color: #4A657E; margin-bottom: 10px; }" +
                "p { color: #64748B; font-size: 14px; line-height: 1.5; margin-bottom: 20px; }" +
                ".btn { display: block; width: 100%; background: #57A773; color: white; border: none; padding: 14px; border-radius: 12px; font-weight: bold; font-size: 16px; margin-bottom: 12px; text-decoration: none; }" +
                ".btn-blue { background: #586F86; }" +
                "</style></head><body>" +
                "<div class='card'>" +
                "<h2>Bodega Central</h2>" +
                "<p>No se pudo conectar al servidor de bodega. Comprueba tu conexión a la red o intenta con el enlace alternativo.</p>" +
                "<a href='" + PRIMARY_URL + "' class='btn'>📶 Reintentar Wi-Fi Bodega</a>" +
                "<a href='" + FALLBACK_URL + "' class='btn btn-blue'>🌐 Intentar por Datos Móviles (4G/5G)</a>" +
                "</div></body></html>";
        view.loadDataWithBaseURL(null, errorHtml, "text/html", "UTF-8", null);
    }
}
