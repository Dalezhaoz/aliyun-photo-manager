from dataclasses import dataclass
from pathlib import Path
from threading import Event
from typing import Callable, Iterable, List, Optional

from .config import OssConfig


IMAGE_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".webp",
    ".heic",
    ".heif",
}


@dataclass
class BrowserEntry:
    key: str
    entry_type: str
    display_name: str


@dataclass
class DownloadResult:
    total_found: int
    downloaded_count: int
    skipped_existing_count: int


def is_photo_key(key: str) -> bool:
    return Path(key).suffix.lower() in IMAGE_EXTENSIONS


def normalize_prefix(prefix: str) -> str:
    normalized_prefix = prefix.strip()
    if normalized_prefix and not normalized_prefix.endswith("/"):
        normalized_prefix = normalized_prefix + "/"
    return normalized_prefix


def build_local_relative_path(object_key: str, prefix: str) -> Path:
    normalized_prefix = normalize_prefix(prefix)
    if normalized_prefix and object_key.startswith(normalized_prefix):
        relative_key = object_key[len(normalized_prefix):]
    else:
        relative_key = object_key
    relative_key = relative_key.lstrip("/")
    if not relative_key:
        relative_key = Path(object_key).name
    return Path(relative_key)


def _detect_provider(config: OssConfig) -> str:
    return config.cloud_type


def _build_aliyun_bucket(config: OssConfig):
    try:
        import oss2
    except ImportError as exc:
        raise ImportError("缺少依赖 oss2，请先执行 `pip install -r requirements.txt`。") from exc

    auth = oss2.Auth(config.access_key_id, config.access_key_secret)
    return oss2.Bucket(auth, config.endpoint, config.bucket_name)


def _extract_cos_region(endpoint: str) -> str:
    cleaned = endpoint.strip()
    if not cleaned:
        raise ValueError("腾讯云 COS 需要填写 Region 或 Endpoint。")
    if cleaned.startswith("http://"):
        cleaned = cleaned[len("http://") :]
    elif cleaned.startswith("https://"):
        cleaned = cleaned[len("https://") :]
    cleaned = cleaned.split("/")[0]
    if cleaned.startswith("cos."):
        cleaned = cleaned[len("cos.") :]
    if cleaned.endswith(".myqcloud.com"):
        cleaned = cleaned[: -len(".myqcloud.com")]
    return cleaned


def _build_tencent_client(config: OssConfig):
    try:
        from qcloud_cos import CosConfig, CosS3Client
    except ImportError as exc:
        raise ImportError(
            "缺少依赖 cos-python-sdk-v5，请先执行 `pip install -r requirements.txt`。"
        ) from exc

    region = _extract_cos_region(config.endpoint)
    client_config = CosConfig(
        Region=region,
        SecretId=config.access_key_id,
        SecretKey=config.access_key_secret,
        Scheme="https",
    )
    return CosS3Client(client_config)


def list_buckets(
    access_key_id: str,
    access_key_secret: str,
    endpoint: str,
    cloud_type: str = "aliyun",
) -> List[str]:
    if cloud_type == "aliyun":
        try:
            import oss2
        except ImportError as exc:
            raise ImportError("缺少依赖 oss2，请先执行 `pip install -r requirements.txt`。") from exc

        auth = oss2.Auth(access_key_id, access_key_secret)
        service = oss2.Service(auth, endpoint)
        return [bucket.name for bucket in service.list_buckets().buckets]

    try:
        from qcloud_cos import CosConfig, CosS3Client
    except ImportError as exc:
        raise ImportError(
            "缺少依赖 cos-python-sdk-v5，请先执行 `pip install -r requirements.txt`。"
        ) from exc

    region = _extract_cos_region(endpoint)
    client = CosS3Client(
        CosConfig(
            Region=region,
            SecretId=access_key_id,
            SecretKey=access_key_secret,
            Scheme="https",
        )
    )
    response = client.list_buckets()
    buckets = response.get("Buckets", {}).get("Bucket", [])
    return [bucket.get("Name", "") for bucket in buckets if bucket.get("Name")]


def _iter_tencent_objects(config: OssConfig, prefix: str = "", delimiter: Optional[str] = None):
    client = _build_tencent_client(config)
    marker = ""
    normalized_prefix = normalize_prefix(prefix) if delimiter == "/" else prefix.strip()
    while True:
        kwargs = {
            "Bucket": config.bucket_name,
            "Prefix": normalized_prefix,
            "MaxKeys": 1000,
        }
        if delimiter:
            kwargs["Delimiter"] = delimiter
        if marker:
            kwargs["Marker"] = marker
        response = client.list_objects(**kwargs)
        yield response
        is_truncated = str(response.get("IsTruncated", "false")).lower() == "true"
        if not is_truncated:
            break
        marker = response.get("NextMarker") or ""
        if not marker:
            break


def list_folder_prefixes(config: OssConfig, prefix: str = "") -> List[str]:
    provider = _detect_provider(config)
    normalized_prefix = normalize_prefix(prefix)

    if provider == "aliyun":
        bucket = _build_aliyun_bucket(config)
        folders: List[str] = []
        result = bucket.list_objects(prefix=normalized_prefix, delimiter="/")
        for folder in result.prefix_list:
            folders.append(folder)
        return folders

    folders: List[str] = []
    for response in _iter_tencent_objects(config, prefix=normalized_prefix, delimiter="/"):
        for folder in response.get("CommonPrefixes", []):
            folder_prefix = folder.get("Prefix", "")
            if folder_prefix:
                folders.append(folder_prefix)
    return folders


def list_browser_entries(config: OssConfig, prefix: str = "") -> List[BrowserEntry]:
    provider = _detect_provider(config)
    normalized_prefix = normalize_prefix(prefix)
    entries: List[BrowserEntry] = []

    if provider == "aliyun":
        bucket = _build_aliyun_bucket(config)
        result = bucket.list_objects(prefix=normalized_prefix, delimiter="/")

        for folder in result.prefix_list:
            entries.append(
                BrowserEntry(
                    key=folder,
                    entry_type="folder",
                    display_name=folder.rstrip("/").split("/")[-1] or "/",
                )
            )

        for obj in result.object_list:
            if obj.key == normalized_prefix:
                continue
            relative = build_local_relative_path(obj.key, normalized_prefix)
            if len(relative.parts) != 1:
                continue
            entries.append(
                BrowserEntry(
                    key=obj.key,
                    entry_type="file",
                    display_name=relative.name,
                )
            )
        return entries

    for response in _iter_tencent_objects(config, prefix=normalized_prefix, delimiter="/"):
        for folder in response.get("CommonPrefixes", []):
            folder_prefix = folder.get("Prefix", "")
            if not folder_prefix:
                continue
            entries.append(
                BrowserEntry(
                    key=folder_prefix,
                    entry_type="folder",
                    display_name=folder_prefix.rstrip("/").split("/")[-1] or "/",
                )
            )

        for obj in response.get("Contents", []):
            key = obj.get("Key", "")
            if not key or key == normalized_prefix:
                continue
            relative = build_local_relative_path(key, normalized_prefix)
            if len(relative.parts) != 1:
                continue
            entries.append(
                BrowserEntry(
                    key=key,
                    entry_type="file",
                    display_name=relative.name,
                )
            )
    return entries


def _iter_object_keys(config: OssConfig, prefix: str) -> Iterable[str]:
    provider = _detect_provider(config)
    normalized_prefix = normalize_prefix(prefix)

    if provider == "aliyun":
        try:
            import oss2
        except ImportError as exc:
            raise ImportError("缺少依赖 oss2，请先执行 `pip install -r requirements.txt`。") from exc

        bucket = _build_aliyun_bucket(config)
        for obj in oss2.ObjectIteratorV2(bucket, prefix=normalized_prefix):
            if obj.is_prefix():
                continue
            yield obj.key
        return

    for response in _iter_tencent_objects(config, prefix=normalized_prefix):
        for obj in response.get("Contents", []):
            key = obj.get("Key", "")
            if key:
                yield key


def count_photos_in_prefix(config: OssConfig, prefix: str) -> int:
    count = 0
    for key in _iter_object_keys(config, prefix):
        if is_photo_key(key):
            count += 1
    return count


def search_objects(
    config: OssConfig,
    keyword: str,
    prefix: str = "",
    max_results: int = 200,
) -> List[str]:
    cleaned_keyword = keyword.strip().lower()
    if not cleaned_keyword:
        return []

    results: List[str] = []
    for key in _iter_object_keys(config, prefix):
        filename = Path(key).name.lower()
        if cleaned_keyword not in filename:
            continue
        results.append(key)
        if len(results) >= max_results:
            break
    return results


def download_photos(
    config: OssConfig,
    prefix: str,
    download_dir: Path,
    dry_run: bool = False,
    skip_existing: bool = True,
    logger: Optional[Callable[[str], None]] = None,
    progress_callback: Optional[Callable[[str, int, int, str], None]] = None,
    cancel_event: Optional[Event] = None,
) -> DownloadResult:
    return download_objects(
        config=config,
        prefix=prefix,
        download_dir=download_dir,
        dry_run=dry_run,
        skip_existing=skip_existing,
        logger=logger,
        progress_callback=progress_callback,
        cancel_event=cancel_event,
        file_filter=is_photo_key,
        stage="download",
    )


def _download_aliyun_object(config: OssConfig, object_key: str, local_path: Path) -> None:
    bucket = _build_aliyun_bucket(config)
    bucket.get_object_to_file(object_key, str(local_path))


def _download_tencent_object(config: OssConfig, object_key: str, local_path: Path) -> None:
    client = _build_tencent_client(config)
    response = client.get_object(Bucket=config.bucket_name, Key=object_key)
    body = response["Body"]
    local_path.write_bytes(body.get_raw_stream().read())


def download_objects(
    config: OssConfig,
    prefix: str,
    download_dir: Path,
    dry_run: bool = False,
    skip_existing: bool = True,
    logger: Optional[Callable[[str], None]] = None,
    progress_callback: Optional[Callable[[str, int, int, str], None]] = None,
    cancel_event: Optional[Event] = None,
    file_filter: Optional[Callable[[str], bool]] = None,
    key_filter: Optional[Callable[[str], bool]] = None,
    stage: str = "download",
) -> DownloadResult:
    download_dir.mkdir(parents=True, exist_ok=True)
    object_keys: List[str] = []
    skipped_existing_count = 0
    downloaded_count = 0

    def log(message: str) -> None:
        if logger is not None:
            logger(message)
        else:
            print(message)

    for key in _iter_object_keys(config, prefix):
        if cancel_event is not None and cancel_event.is_set():
            break
        if key_filter is not None and not key_filter(key):
            continue
        if file_filter is not None and not file_filter(key):
            continue
        object_keys.append(key)

    total = len(object_keys)
    if progress_callback is not None:
        progress_callback(stage, 0, total, "")
    log(f"共找到 {total} 个可下载文件。")

    for index, object_key in enumerate(object_keys, start=1):
        if cancel_event is not None and cancel_event.is_set():
            log(f"下载已取消，已处理 {index - 1}/{total} 个文件。")
            if progress_callback is not None:
                progress_callback(stage, index - 1, total, "")
            break

        relative_path = build_local_relative_path(object_key, prefix)
        local_path = download_dir / relative_path

        if skip_existing and local_path.exists():
            if index == 1 or index % 100 == 0 or index == total:
                log(f"下载进度 {index}/{total}，跳过已存在文件。")
            skipped_existing_count += 1
            if progress_callback is not None:
                progress_callback(stage, index, total, object_key)
            continue

        if index == 1 or index % 100 == 0 or index == total:
            log(f"下载进度 {index}/{total}：{Path(object_key).name}")

        if not dry_run:
            local_path.parent.mkdir(parents=True, exist_ok=True)
            if _detect_provider(config) == "aliyun":
                _download_aliyun_object(config, object_key, local_path)
            else:
                _download_tencent_object(config, object_key, local_path)
        downloaded_count += 1

        if progress_callback is not None:
            progress_callback(stage, index, total, object_key)

    return DownloadResult(
        total_found=total,
        downloaded_count=downloaded_count,
        skipped_existing_count=skipped_existing_count,
    )
