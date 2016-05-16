package com.whh.utils;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.log4j.Logger;

import javax.net.ssl.*;
import java.io.ByteArrayOutputStream;
import java.io.DataOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.*;
import java.nio.charset.Charset;
import java.security.KeyManagementException;
import java.security.NoSuchAlgorithmException;
import java.security.cert.CertificateException;
import java.security.cert.X509Certificate;
import java.util.HashMap;
import java.util.Map;

/**
 * 需要使用gosn转换
 * Created by xuzhuo on 2015/12/14.
 * Http请求Utils
 */
public class HttpClientUtils {
    private static final Logger logger = Logger.getLogger(HttpClientUtils.class);
    private static final String ContentType = "Content-Type";
    private static final String FORM_URLENCODED = "application/x-www-form-urlencoded";
    private static final String JSON_TYPE = "application/json";

    private static Gson gson = new GsonBuilder().create();

    /**
     * POST请求 application/x-www-form-urlencoded类型
     *
     * @param url    请求URL
     * @param params 请求key
     * @param data   请求值
     * @return 返回数据, 请求出错返回null
     */
    public static String post(String url, String params, String data) {
        String result = null;
        try {
            URLConnection conn = createURLConnection(url);
            result = post(conn, params + "=" + data, ContentType, FORM_URLENCODED);
        } catch (NoSuchAlgorithmException | KeyManagementException | URISyntaxException | IOException e) {
            logger.error("请求出错--->URL: " + url + ";\t" + "params:" + params + ";\t" + "data:" + data);
            e.printStackTrace();
        }
        return result;
    }

    /**
     * POST请求 application/x-www-form-urlencoded类型
     *
     * @param url    请求URL
     * @param params 请求key
     * @param data   请求值
     * @return 请求值
     */
    public static String post(String url, String params, Map<String, ?> data) {
        return post(url, params, gson.toJson(data));
    }

    /**
     * POST请求 application/x-www-form-urlencoded类型
     *
     * @param url  请求URL
     * @param data 请求数据
     * @return
     */
    public static String post(String url, Map<String, ?> data) {
        String result = null;
        StringBuilder params = new StringBuilder();
        for (Map.Entry<String, ?> entry : data.entrySet()) {
            params.append(entry.getKey()).append("=").append(entry.getValue()).append("&");
        }
        String paramsStr = params.toString();
        if (paramsStr.endsWith("&")){
            paramsStr = paramsStr.substring(0, paramsStr.length() - 1);
        }
        try {
            result = post(createURLConnection(url), paramsStr, ContentType, FORM_URLENCODED);
        } catch (NoSuchAlgorithmException | KeyManagementException | URISyntaxException | IOException e) {
            logger.error("请求出错--->URL: " + url + ";\t" + "data:" + params.toString());
            e.printStackTrace();
        }
        return result;
    }

    /**
     * POST请求 application/json
     *
     * @param url  请求URL
     * @param json 请求json数据
     * @return 返回数据, 请求出错返回null
     */
    public static String postJSON(String url, String json) {
        String result = null;
        try {
            URLConnection conn = createURLConnection(url);
            result = post(conn, json, ContentType, JSON_TYPE);
        } catch (NoSuchAlgorithmException | KeyManagementException | URISyntaxException | IOException e) {
            logger.error("请求出错--->URL: " + url + ";\t" + "json:" + json);
            e.printStackTrace();
        }
        return result;
    }

    /**
     * POST请求 application/json
     *
     * @param url    请求URL
     * @param params 请求参数Map类型
     * @return 返回数据, 请求出错返回null
     */
    public static String postJSON(String url, Map<String, ?> params) {
        return postJSON(url, gson.toJson(params));
    }

    /**
     * POST请求封装Map application/json
     *
     * @param url   请求URL
     * @param key   请求参数Key
     * @param value 请求参数Value
     * @return 返回数据, 请求出错返回null
     */
    public static String postJSON(String url, String key, Object value) {
        Map<String, Object> params = new HashMap<>();
        params.put(key, value);
        return postJSON(url, params);
    }

    /**
     * GET 请求
     *
     * @param conn URL连接
     * @return 返回数据
     * @throws URISyntaxException
     * @throws IOException
     */
    private static String get(URLConnection conn) throws URISyntaxException, IOException {
        String url = conn.getURL().toString();
        logger.error("URL: " + url);
        // 设置通用的请求属性
        conn.setRequestProperty("Accept", "*/*");
        conn.setRequestProperty("Connection", "keep-alive");
        conn.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.134 Safari/537.36");
        conn.setConnectTimeout(30000);
        conn.setReadTimeout(30000);
        conn.connect();
        // 定义BufferedReader输入流来读取URL的响应
        String result = null;
        InputStream is = conn.getInputStream();
        // 读取返回数据
        if (is != null) {
            result = getHttpResponseToStr(is);
        }
        logger.error("URL: " + url + ";\t" + "return:" + result);
        return result;
    }

    /**
     * 获取返回body
     * @param is 流
     * @return 返回body
     * @throws IOException
     */
    private static String getHttpResponseToStr(InputStream is) throws IOException {
        String result;ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        byte[] buffer = new byte[1024];
        int len;
        while ((len = is.read(buffer)) != -1) {
            outStream.write(buffer, 0, len);
        }
        is.close();
        result = new String(outStream.toByteArray(), Charset.forName("utf-8"));
        return result;
    }

    /**
     * GET 请求
     *
     * @param url 请求URL
     * @return 返回数据
     */
    public static String get(String url) {
        String result = null;
        try {
            result = get(createURLConnection(url));
        } catch (IOException | NoSuchAlgorithmException | KeyManagementException | URISyntaxException e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * GET 请求
     *
     * @param url    请求URL
     * @param params 请求数据
     * @return 返回数据
     */
    public static String get(String url, String params) {
        return get(url + "?" + params);
    }

    /**
     * GET 请求
     *
     * @param url  请求URL
     * @param key  请求数据KEY
     * @param data 请求数据VAL
     * @return 返回数据
     */
    public static String get(String url, String key, String data) {
        return get(url, key + "=" + data);
    }

    /**
     * GET 请求
     *
     * @param url    请求URL
     * @param params 请求数据
     * @return 返回数据
     */
    public static String get(String url, Map<String, ?> params) {
        StringBuilder urlParams = new StringBuilder();
        for (Map.Entry<String, ?> entry : params.entrySet()) {
            urlParams.append(entry.getKey()).append("=").append(entry.getValue()).append("&");
        }
        return get(url, urlParams.substring(0, urlParams.length() - 1));
    }

    /**
     * POST封装
     *
     * @param conn        URLConnection
     * @param data        请求数据
     * @param headerKey   headerKey
     * @param headerValue headerValue
     * @return 返回数据
     * @throws NoSuchAlgorithmException
     * @throws KeyManagementException
     * @throws IOException
     * @throws URISyntaxException
     */
    private static String post(URLConnection conn, String data, String headerKey, String headerValue) throws NoSuchAlgorithmException, KeyManagementException, IOException, URISyntaxException {
        Map<String, String> header = new HashMap<>();
        header.put(headerKey, headerValue);
        return post(conn, data, header);
    }

    /**
     * 实际请求
     *
     * @param conn   url连接
     * @param data   请求数据
     * @param header 请求header
     * @return 返回数据
     * @throws URISyntaxException
     * @throws IOException
     */
    private static String post(URLConnection conn, String data, Map<String, String> header) throws URISyntaxException, IOException {
        String url = conn.getURL().toString();
        logger.error("URL: " + url + ";\t" + "data:" + data);
        // 设置通用的请求属性
        conn.setRequestProperty("Accept", "*/*");
        conn.setRequestProperty("Connection", "keep-alive");
        conn.setRequestProperty("User-Agent", "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.134 Safari/537.36");
        conn.setConnectTimeout(30000);
        conn.setReadTimeout(30000);
        for (Map.Entry<String, String> headerEntry : header.entrySet()) {
            conn.setRequestProperty(headerEntry.getKey(), headerEntry.getValue());
        }
        // 发送POST请求必须设置如下两行
        conn.setDoOutput(true);
        conn.setDoInput(true);
        // 获取URLConnection对象对应的输出流
        DataOutputStream out = new DataOutputStream(conn.getOutputStream());
        out.write(data.getBytes("utf-8"));
        out.flush();
        out.close();
        // 定义BufferedReader输入流来读取URL的响应
        String result = null;
        InputStream is = conn.getInputStream();
        // 读取返回数据
        if (is != null) {
            result = getHttpResponseToStr(is);
        }
        logger.error("URL: " + url + ";\t" + "data:" + data + ";\t" + "return:" + result);
        return result;
    }

    /**
     * 获取URLConnection
     *
     * @param url 请求URL,依据URL判断生成HttpsURLConnection还是普通URLConnection
     * @return URLConnection
     * @throws IOException
     * @throws NoSuchAlgorithmException
     * @throws KeyManagementException
     */
    private static URLConnection createURLConnection(String url) throws IOException, NoSuchAlgorithmException, KeyManagementException, URISyntaxException {
        URL realUrl = new URL(url);
        URI uri = new URI(realUrl.getProtocol(), realUrl.getUserInfo(), realUrl.getHost(), realUrl.getPort(), realUrl.getPath(), realUrl.getQuery(), realUrl.getRef());
        realUrl = uri.toURL();
        if (url.matches("https://.*?")) {
            SSLContext sc = SSLContext.getInstance("SSL");
            sc.init(null, new TrustManager[]{new TrustAnyTrustManager()}, new java.security.SecureRandom());
            HttpsURLConnection httpsConn = (HttpsURLConnection) realUrl.openConnection();
            httpsConn.setSSLSocketFactory(sc.getSocketFactory());
            httpsConn.setHostnameVerifier(new TrustAnyHostnameVerifier());
            return httpsConn;
        } else {
            return realUrl.openConnection();
        }
    }

    private static class TrustAnyTrustManager implements X509TrustManager {

        public void checkClientTrusted(X509Certificate[] chain, String authType)
                throws CertificateException {
        }

        public void checkServerTrusted(X509Certificate[] chain, String authType)
                throws CertificateException {
        }

        public X509Certificate[] getAcceptedIssuers() {
            return new X509Certificate[]{};
        }
    }

    private static class TrustAnyHostnameVerifier implements HostnameVerifier {
        public boolean verify(String hostname, SSLSession session) {
            return true;
        }
    }
}
