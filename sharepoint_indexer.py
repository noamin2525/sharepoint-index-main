import requests
import os
from urllib.parse import quote, unquote
import json
from flask import Flask, render_template_string, send_file, jsonify, request, redirect, Response, stream_with_context
import io
from urllib.parse import quote
app = Flask(__name__)


class SharePointIndexer :
    def __init__(self) :
        self.client_id = "9da3c953-1cd2-47b2-9905-a35584254d7a"
        self.client_secret = "yiu8Q~I2VaOFlDUqSSF~RloOEFZIKCULUzD4wan-"
        self.tenant_id = "53729536-9ac2-4744-95f3-56410884a077"
        self.sharepoint_folder = "/israeli-movies-series"
        self.sharepoint_site = "5c8hk2.sharepoint.com:/sites/movies"

        self.access_token = None
        self.site_id = None
        self.drive_id = None

    def get_access_token(self) :
        """Get OAuth access token from Azure AD"""
        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        data = {
            'client_id' : self.client_id,
            'client_secret' : self.client_secret,
            'scope' : 'https://graph.microsoft.com/.default',
            'grant_type' : 'client_credentials'
        }

        response = requests.post(url, data=data)
        response.raise_for_status()

        self.access_token = response.json()['access_token']
        return self.access_token

    def get_site_id(self) :
        """Get the SharePoint site ID"""
        headers = {
            'Authorization' : f'Bearer {self.access_token}',
            'Accept' : 'application/json'
        }

        url = f"https://graph.microsoft.com/v1.0/sites/{self.sharepoint_site}"

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        self.site_id = response.json()['id']
        return self.site_id

    def get_drive_id(self) :
        """Get the default document library drive ID"""
        if self.drive_id :
            return self.drive_id

        headers = {
            'Authorization' : f'Bearer {self.access_token}',
            'Accept' : 'application/json'
        }

        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drive"

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        self.drive_id = response.json()['id']
        return self.drive_id

    def list_files(self, folder_path='') :
        """List all files in the specified folder"""
        headers = {
            'Authorization' : f'Bearer {self.access_token}',
            'Accept' : 'application/json'
        }

        drive_id = self.get_drive_id()

        if folder_path :
            encoded_path = quote(folder_path)
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/root:/{encoded_path}:/children"
        else :
            url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/root/children"

        all_items = []

        while url :
            response = requests.get(url, headers=headers)
            response.raise_for_status()

            data = response.json()
            all_items.extend(data.get('value', []))

            url = data.get('@odata.nextLink')

        return all_items

    def download_file_stream(self, file_id) :
        """Download a file by its ID and return as stream"""
        headers = {
            'Authorization' : f'Bearer {self.access_token}'
        }

        drive_id = self.get_drive_id()
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/items/{file_id}/content"

        # allow caller to pass through extra headers (eg Range)
        extra_headers = getattr(self, '_extra_headers', None)
        if extra_headers:
            headers.update(extra_headers)

        response = requests.get(url, headers=headers, stream=True)
        response.raise_for_status()

        return response

    def get_download_url(self, file_id):
        """Retrieve the pre-authenticated download URL from Microsoft Graph for the item."""
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Accept': 'application/json'
        }

        drive_id = self.get_drive_id()
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{drive_id}/items/{file_id}"
        params = {
            '$select': '@microsoft.graph.downloadUrl'
        }
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()
        return data.get('@microsoft.graph.downloadUrl')


# Initialize indexer
indexer = SharePointIndexer()

# HTML Template
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint File Browser</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }

        .header h1 {
            font-size: 2em;
            margin-bottom: 10px;
        }

        .breadcrumb {
            background: #f8f9fa;
            padding: 15px 30px;
            border-bottom: 1px solid #dee2e6;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .breadcrumb a {
            color: #667eea;
            text-decoration: none;
            font-weight: 500;
        }

        .breadcrumb a:hover {
            text-decoration: underline;
        }

        .breadcrumb span {
            color: #6c757d;
        }

        .content {
            padding: 30px;
        }

        .item-list {
            list-style: none;
        }

        .item {
            display: flex;
            align-items: center;
            padding: 15px;
            border-bottom: 1px solid #e9ecef;
            transition: background 0.2s;
        }

        .item:hover {
            background: #f8f9fa;
        }

        .item-icon {
            font-size: 24px;
            margin-right: 15px;
            width: 30px;
            text-align: center;
        }

        .item-info {
            flex: 1;
        }

        .item-name {
            font-weight: 500;
            color: #212529;
            margin-bottom: 5px;
        }

        .item-meta {
            font-size: 0.85em;
            color: #6c757d;
        }

        .item-actions {
            display: flex;
            gap: 10px;
        }

        .btn {
            padding: 8px 16px;
            border: none;
            border-radius: 6px;
            cursor: pointer;
            font-size: 0.9em;
            text-decoration: none;
            transition: all 0.2s;
        }

        .btn-primary {
            background: #667eea;
            color: white;
        }

        .btn-primary:hover {
            background: #5568d3;
        }

        .btn-secondary {
            background: #6c757d;
            color: white;
        }

        .btn-secondary:hover {
            background: #5a6268;
        }

        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #6c757d;
        }

        .empty-state-icon {
            font-size: 64px;
            margin-bottom: 20px;
        }

        .loading {
            text-align: center;
            padding: 40px;
            color: #6c757d;
        }

        .spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📁 SharePoint File Browser</h1>
            <p>Browse and download your files</p>
        </div>

        <div class="breadcrumb">
            <a href="/">🏠 Home</a>
            {% if path %}
                {% set parts = path.split('/') %}
                {% for part in parts %}
                    {% if part %}
                        <span>/</span>
                        <a href="/?path={{ parts[:loop.index]|join('/') }}">{{ part }}</a>
                    {% endif %}
                {% endfor %}
            {% endif %}
        </div>

        <div class="content">
            <div id="loading" class="loading">
                <div class="spinner"></div>
                <p>Loading files...</p>
            </div>

            <ul id="item-list" class="item-list" style="display: none;">
            </ul>

            <div id="empty-state" class="empty-state" style="display: none;">
                <div class="empty-state-icon">📭</div>
                <h3>No files found</h3>
                <p>This folder is empty</p>
            </div>
        </div>
    </div>

    <script>
        const currentPath = new URLSearchParams(window.location.search).get('path') || '';

        async function loadFiles() {
            try {
                const response = await fetch(`/api/list?path=${encodeURIComponent(currentPath)}`);
                const data = await response.json();

                document.getElementById('loading').style.display = 'none';

                if (data.folders.length === 0 && data.files.length === 0) {
                    document.getElementById('empty-state').style.display = 'block';
                    return;
                }

                const itemList = document.getElementById('item-list');
                itemList.style.display = 'block';
                itemList.innerHTML = '';

                // Add folders first
                data.folders.forEach(folder => {
                    const li = document.createElement('li');
                    li.className = 'item';
                    li.innerHTML = `
                        <div class="item-icon">📁</div>
                        <div class="item-info">
                            <div class="item-name">${escapeHtml(folder.name)}</div>
                            <div class="item-meta">Modified: ${new Date(folder.modified).toLocaleDateString()}</div>
                        </div>
                        <div class="item-actions">
                            <a href="/?path=${encodeURIComponent(folder.full_path)}" class="btn btn-primary">Open</a>
                        </div>
                    `;
                    itemList.appendChild(li);
                });

                // Add files
                data.files.forEach(file => {
                    const sizeMB = (file.size / (1024 * 1024)).toFixed(2);
                    const li = document.createElement('li');
                    li.className = 'item';
                    li.innerHTML = `
                        <div class="item-icon">📄</div>
                        <div class="item-info">
                            <div class="item-name">${escapeHtml(file.name)}</div>
                            <div class="item-meta">${sizeMB} MB • Modified: ${new Date(file.modified).toLocaleDateString()}</div>
                        </div>
                        <div class="item-actions">
                            <a href="${file.download_url}" class="btn btn-primary" download>Download</a>
                            <a href="${file.stream_url}" class="btn btn-secondary" target="_blank">Stream</a>
                            <a href="${file.strm_url}" class="btn btn-secondary" target="_blank">.strm</a>
                            <a href="${file.m3u_url}" class="btn btn-secondary" target="_blank">.m3u</a>
                        </div>
                    `;
                    itemList.appendChild(li);
                });

            } catch (error) {
                document.getElementById('loading').innerHTML = `
                    <p style="color: red;">Error loading files: ${error.message}</p>
                `;
            }
        }

        function escapeHtml(text) {
            const div = document.createElement('div');
            div.textContent = text;
            return div.innerHTML;
        }

        loadFiles();
    </script>
</body>
</html>
"""


@app.route('/')
def index() :
    path = request.args.get('path', '')
    return render_template_string(HTML_TEMPLATE, path=path)


@app.route('/api/list')
def list_files() :
    try:
        # Initialize on each request for serverless
        if not indexer.access_token:
            indexer.get_access_token()
            indexer.get_site_id()
            indexer.get_drive_id()
        
        path = request.args.get('path', '')

        # Combine base folder with requested path
        full_path = indexer.sharepoint_folder
        if path :
            full_path = f"{indexer.sharepoint_folder}/{path}".replace('//', '/')

        items = indexer.list_files(full_path)

        result = {
            'folders' : [],
            'files' : []
        }

        for item in items :
            if 'folder' in item :
                folder_path = f"{path}/{item['name']}" if path else item['name']
                result['folders'].append({
                    'name' : item['name'],
                    'full_path' : folder_path,
                    'id' : item['id'],
                    'created' : item['createdDateTime'],
                    'modified' : item['lastModifiedDateTime']
                })
            elif 'file' in item :
                # Construct local endpoints for download and streaming
                # /download?file_id=...&filename=...
                local_download = f"/download?file_id={item['id']}&filename={quote(item['name'])}"
                local_stream = f"/stream?file_id={item['id']}&filename={quote(item['name'])}"
                # .strm file (Stremio accepts simple URLs or .strm files containing the direct link)
                local_strm = f"/strm?file_id={item['id']}&filename={quote(item['name'])}"
                # .m3u entry for IPTV VOD (can be downloaded and used in IPTV players)
                local_m3u = f"/m3u?file_id={item['id']}&filename={quote(item['name'])}"

                result['files'].append({
                    'name' : item['name'],
                    'id' : item['id'],
                    'size' : item['size'],
                    'download_url' : local_download,
                    'stream_url' : local_stream,
                    'strm_url' : local_strm,
                    'm3u_url' : local_m3u,
                    'created' : item['createdDateTime'],
                    'modified' : item['lastModifiedDateTime']
                })

        return jsonify(result)
    except Exception as e :
        return jsonify({'error' : str(e)}), 500


@app.route('/download')
def download_file() :
    try :
        file_id = request.args.get('file_id')
        filename = request.args.get('filename')

        print(f"Download request - File ID: {file_id}, Filename: {filename}")

        if not file_id or not filename :
            return "Missing file_id or filename parameter", 400

        response = indexer.download_file_stream(file_id)

        # Stream the file to the client
        return send_file(
            io.BytesIO(response.content),
            as_attachment=True,
            download_name=filename,
            mimetype='application/octet-stream'
        )
    except Exception as e :
        print(f"Download error: {str(e)}")
        return f"Error downloading file: {str(e)}", 500


@app.route('/stream')
def stream_file():
    """Stream file proxy that supports Range requests for seeking (use in VLC/IPTV players)."""
    try:
        file_id = request.args.get('file_id')
        filename = request.args.get('filename', 'stream')

        if not file_id:
            return "Missing file_id parameter", 400

        # Rather than proxying the full media through the serverless function (which can fail),
        # obtain the direct pre-authenticated download URL from Graph and redirect the client to it.
        # This allows VLC and IPTV players to perform Range requests directly against SharePoint's URL.
        download_url = indexer.get_download_url(file_id)
        if not download_url:
            return "Could not retrieve download URL", 500

        # Redirect client to the direct URL
        return redirect(download_url)

    except Exception as e:
        print(f"Stream error: {e}")
        return f"Error streaming file: {str(e)}", 500


@app.route('/strm')
def serve_strm():
    """Serve a .strm file that points to the /stream endpoint. Useful for Stremio or clients that accept .strm files."""
    file_id = request.args.get('file_id')
    filename = request.args.get('filename', 'stream')
    if not file_id:
        return "Missing file_id parameter", 400

    # Provide a .strm file that points directly to the Graph download URL
    download_url = indexer.get_download_url(file_id)
    if not download_url:
        return "Could not retrieve download URL", 500

    content = download_url
    resp = Response(content, mimetype='application/x-mpegURL')
    resp.headers['Content-Disposition'] = f'attachment; filename="{filename}.strm"'
    return resp


@app.route('/m3u')
def serve_m3u():
    """Serve a tiny M3U playlist entry pointing to the /stream endpoint for IPTV VOD players."""
    file_id = request.args.get('file_id')
    filename = request.args.get('filename', 'stream')
    if not file_id:
        return "Missing file_id parameter", 400

    download_url = indexer.get_download_url(file_id)
    if not download_url:
        return "Could not retrieve download URL", 500

    # Basic M3U with a single entry pointing to the direct download URL
    m3u = f"#EXTM3U\n#EXTINF:-1,{filename}\n{download_url}\n"
    resp = Response(m3u, mimetype='application/x-mpegurl')
    resp.headers['Content-Disposition'] = f'attachment; filename="{filename}.m3u"'
    return resp


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
