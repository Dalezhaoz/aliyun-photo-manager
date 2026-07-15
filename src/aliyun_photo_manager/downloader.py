from __future__ import annotations

from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from threading import Event, local
from typing import Callable, Iterable, List, Optional

from .config import OssConfig
from .config import normalize_cloud_type


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

DEFAULT_MAX_WORKERS = {
    "aliyun": 4,
    "tencent": 3,
}
_CLIENT_LOCAL = local()


@dataclass
class BrowserEntry:
    key: str
    entry_type: str
    display_name: str


@dataclass
class BucketInfo:
    name: str
    location: str = ""


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
    # 本地目录尽量保留“前缀之后”的相对结构，方便和云端路径互相对应。
    if normalized_prefix and object_key.startswith(normalized_prefix):
        relative_key = object_key[len(normalized_prefix):]
    else:
        relative_key = object_key
    relative_key = relative_key.lstrip("/")
    if not relative_key:
        relative_key = Path(object_key).name
    # Object names are remote input. Never allow one to escape the selected
    # download directory, especially when the app runs on Windows.
    if "\\" in relative_key or "\x00" in relative_key:
        raise ValueError(f"云端对象路径不安全：{object_key}")
    relative_path = PurePosixPath(relative_key)
    if relative_path.is_absolute() or any(part in {"", ".", ".."} for part in relative_path.parts):
        raise ValueError(f"云端对象路径不安全：{object_key}")
    if any(":" in part for part in relative_path.parts):
        raise ValueError(f"云端对象路径不安全：{object_key}")
    return Path(*relative_path.parts)


def _detect_provider(config: OssConfig) -> str:
    return config.cloud_type


def _build_aliyun_bucket(config: OssConfig):
    try:
        import oss2
    except ImportError as exc:
        raise ImportError("缺少依赖 oss2，请先执行 `pip install -r requirements.txt`。") from exc

    auth = oss2.Auth(config.access_key_id, config.access_key_secret)
    return oss2.Bucket(auth, config.endpoint, config.bucket_name)


def _iter_aliyun_current_level(config: OssConfig, prefix: str):
    try:
        import oss2
    except ImportError as exc:
        raise ImportError("缺少依赖 oss2，请先执行 `pip install -r requirements.txt`。") from exc

    bucket = _build_aliyun_bucket(config)
    for result in oss2.ObjectIteratorV2(bucket, prefix=prefix, delimiter="/", max_keys=1000):
        prefix_list = getattr(result, "prefix_list", None)
        object_list = getattr(result, "object_list", None)
        if prefix_list is not None or object_list is not None:
            for folder in prefix_list or []:
                yield folder, True
            for obj in object_list or []:
                yield getattr(obj, "key", ""), False
            continue

        key = getattr(result, "key", "")
        if key:
            yield key, result.is_prefix()


def _extract_cos_region(endpoint: str) -> str:
    cleaned = endpoint.strip()
    if not cleaned:
        raise ValueError("腾讯云 COS 需要填写 Region 或 Endpoint。")
    # GUI 里既允许直接填 Region，也允许填完整 Endpoint，这里统一折成 Region。
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


def list_bucket_infos(
    access_key_id: str,
    access_key_secret: str,
    endpoint: str,
    cloud_type: str = "aliyun",
) -> List[BucketInfo]:
    cloud_type = normalize_cloud_type(cloud_type)
    if cloud_type == "aliyun":
        try:
            import oss2
        except ImportError as exc:
            raise ImportError("缺少依赖 oss2，请先执行 `pip install -r requirements.txt`。") from exc

        auth = oss2.Auth(access_key_id, access_key_secret)
        service = oss2.Service(auth, endpoint)
        return [
            BucketInfo(name=bucket.name, location=getattr(bucket, "location", "") or "")
            for bucket in service.list_buckets().buckets
        ]

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
    return [
        BucketInfo(name=bucket.get("Name", ""), location=bucket.get("Location", "") or "")
        for bucket in buckets
        if bucket.get("Name")
    ]


def list_buckets(
    access_key_id: str,
    access_key_secret: str,
    endpoint: str,
    cloud_type: str = "aliyun",
) -> List[str]:
    return [
        bucket.name
        for bucket in list_bucket_infos(
            access_key_id=access_key_id,
            access_key_secret=access_key_secret,
            endpoint=endpoint,
            cloud_type=cloud_type,
        )
    ]


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
        folders: List[str] = []
        for key, is_prefix in _iter_aliyun_current_level(config, normalized_prefix):
            if is_prefix:
                folders.append(key)
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
        seen_folders: set[str] = set()
        for key, is_prefix in _iter_aliyun_current_level(config, normalized_prefix):
            if is_prefix:
                if key in seen_folders:
                    continue
                seen_folders.add(key)
                entries.append(
                    BrowserEntry(
                        key=key,
                        entry_type="folder",
                        display_name=key.rstrip("/").split("/")[-1] or "/",
                    )
                )
                continue

            if key == normalized_prefix:
                continue
            relative = build_local_relative_path(key, normalized_prefix)
            # 浏览面板只显示当前层级，深层文件要进入子目录后再看。
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


def search_folder_entries(config: OssConfig, keyword: str) -> List[BrowserEntry]:
    provider = _detect_provider(config)
    search_prefix = keyword.strip().lstrip("/")
    entries: dict[str, BrowserEntry] = {}

    if provider == "aliyun":
        for key, is_prefix in _iter_aliyun_current_level(config, search_prefix):
            if is_prefix:
                entries.setdefault(
                    key,
                    BrowserEntry(
                        key=key,
                        entry_type="folder",
                        display_name=key.rstrip("/").split("/")[-1] or "/",
                    ),
                )
                continue

            if key in {search_prefix, normalize_prefix(search_prefix)}:
                continue
            entries.setdefault(
                key,
                BrowserEntry(
                    key=key,
                    entry_type="file",
                    display_name=key.rstrip("/").split("/")[-1] or key,
                ),
            )
        return [entries[key] for key in sorted(entries)]

    # 腾讯云：直接用关键词作为 Prefix（不加 "/"），匹配所有以关键词开头的对象。
    client = _build_tencent_client(config)
    marker = ""
    while True:
        kwargs = {
            "Bucket": config.bucket_name,
            "Prefix": search_prefix,
            "Delimiter": "/",
            "MaxKeys": 1000,
        }
        if marker:
            kwargs["Marker"] = marker
        response = client.list_objects(**kwargs)

        for folder in response.get("CommonPrefixes", []):
            folder_prefix = folder.get("Prefix", "")
            if not folder_prefix:
                continue
            entries.setdefault(
                folder_prefix,
                BrowserEntry(
                    key=folder_prefix,
                    entry_type="folder",
                    display_name=folder_prefix.rstrip("/").split("/")[-1] or "/",
                ),
            )

        for obj in response.get("Contents", []):
            key = obj.get("Key", "")
            if not key or key in {search_prefix, normalize_prefix(search_prefix)}:
                continue
            entries.setdefault(
                key,
                BrowserEntry(
                    key=key,
                    entry_type="file",
                    display_name=key.rstrip("/").split("/")[-1] or key,
                ),
            )

        is_truncated = str(response.get("IsTruncated", "false")).lower() == "true"
        if not is_truncated:
            break
        marker = response.get("NextMarker") or ""
        if not marker:
            break
    return [entries[key] for key in sorted(entries)]


def _iter_object_keys(config: OssConfig, prefix: str, *, directory_prefix: bool = True) -> Iterable[str]:
    provider = _detect_provider(config)
    normalized_prefix = normalize_prefix(prefix) if directory_prefix else prefix.strip()

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
    key_filter: Optional[Callable[[str], bool]] = None,
    object_prefixes: Optional[Iterable[str]] = None,
    max_workers: int | None = None,
) -> DownloadResult:
    # 照片下载只是通用下载器的一个特化入口：额外按图片后缀过滤。
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
        key_filter=key_filter,
        object_prefixes=object_prefixes,
        max_workers=max_workers,
        stage="download",
    )


def _download_aliyun_object(config: OssConfig, object_key: str, local_path: Path) -> None:
    bucket = getattr(_CLIENT_LOCAL, "aliyun_bucket", None)
    if bucket is None:
        bucket = _build_aliyun_bucket(config)
        _CLIENT_LOCAL.aliyun_bucket = bucket
    bucket.get_object_to_file(object_key, str(local_path))


def _download_tencent_object(config: OssConfig, object_key: str, local_path: Path) -> None:
    client = getattr(_CLIENT_LOCAL, "tencent_client", None)
    if client is None:
        client = _build_tencent_client(config)
        _CLIENT_LOCAL.tencent_client = client
    response = client.get_object(Bucket=config.bucket_name, Key=object_key)
    body = response["Body"]
    local_path.write_bytes(body.get_raw_stream().read())


def _resolve_worker_count(config: OssConfig, total: int | None, requested_workers: int | None = None) -> int:
    provider = _detect_provider(config)
    default = DEFAULT_MAX_WORKERS.get(provider, 3)
    worker_limit = requested_workers if requested_workers is not None and requested_workers > 0 else default
    if total is None:
        return max(1, worker_limit)
    if total <= 1:
        return 1
    return max(1, min(worker_limit, total))


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
    object_prefixes: Optional[Iterable[str]] = None,
    max_workers: int | None = None,
    stage: str = "download",
) -> DownloadResult:
    download_dir.mkdir(parents=True, exist_ok=True)
    skipped_existing_count = 0
    downloaded_count = 0
    total_found = 0
    processed_count = 0
    scanned_count = 0
    scan_completed = False

    def log(message: str) -> None:
        if logger is not None:
            logger(message)
        else:
            print(message)

    worker_count = _resolve_worker_count(config, None, max_workers)

    def worker(object_key: str, local_path: Path) -> str:
        if cancel_event is not None and cancel_event.is_set():
            return object_key
        local_path.parent.mkdir(parents=True, exist_ok=True)
        partial_path = local_path.with_name(f"{local_path.name}.part")
        partial_path.unlink(missing_ok=True)
        try:
            if _detect_provider(config) == "aliyun":
                _download_aliyun_object(config, object_key, partial_path)
            else:
                _download_tencent_object(config, object_key, partial_path)
            partial_path.replace(local_path)
        except Exception:
            partial_path.unlink(missing_ok=True)
            raise
        return object_key

    in_flight: dict = {}

    def report_scan(object_key: str) -> None:
        if progress_callback is not None:
            progress_callback("listing", scanned_count, 0, object_key)

    def report_processed(object_key: str) -> None:
        if progress_callback is not None:
            progress_callback(stage, processed_count, total_found if scan_completed else 0, object_key)

    def process_completed() -> None:
        nonlocal downloaded_count, processed_count
        if not in_flight:
            return
        done, _ = wait(in_flight.keys(), return_when=FIRST_COMPLETED)
        for future in done:
            object_key = in_flight.pop(future)
            future.result()
            downloaded_count += 1
            processed_count += 1
            if processed_count == 1 or processed_count % 100 == 0:
                log(f"下载进度：已处理 {processed_count} 个文件。")
            report_processed(object_key)

    def iter_selected_objects() -> Iterable[str]:
        if object_prefixes is None:
            yield from _iter_object_keys(config, prefix)
            return

        base_prefix = normalize_prefix(prefix)
        seen_keys: set[str] = set()
        for object_prefix in object_prefixes:
            cleaned_object_prefix = object_prefix.strip().lstrip("/")
            if not cleaned_object_prefix:
                continue
            # OSS/COS 的 Prefix 不需要以 / 结束；这里直接按文件名前缀检索。
            lookup_prefix = f"{base_prefix}{cleaned_object_prefix}"
            for object_key in _iter_object_keys(config, lookup_prefix, directory_prefix=False):
                if object_key in seen_keys:
                    continue
                seen_keys.add(object_key)
                yield object_key

    with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix="bucket-download") as executor:
        for object_key in iter_selected_objects():
            if cancel_event is not None and cancel_event.is_set():
                break
            scanned_count += 1
            if scanned_count == 1 or scanned_count % 500 == 0:
                report_scan(object_key)
                log(f"正在扫描云端目录：已检查 {scanned_count} 个对象，匹配到 {total_found} 个文件。")
            # 证件资料“按名单下载”和照片“按文件名前缀下载”都复用这层过滤。
            if key_filter is not None and not key_filter(object_key):
                continue
            if file_filter is not None and not file_filter(object_key):
                continue

            total_found += 1
            if dry_run:
                downloaded_count += 1
                processed_count += 1
                report_processed(object_key)
                continue

            relative_path = build_local_relative_path(object_key, prefix)
            local_path = download_dir / relative_path
            if skip_existing and local_path.exists():
                skipped_existing_count += 1
                processed_count += 1
                report_processed(object_key)
                continue

            while len(in_flight) >= worker_count:
                process_completed()
            future = executor.submit(worker, object_key, local_path)
            in_flight[future] = object_key

        scan_completed = True
        log(f"云端扫描完成：共找到 {total_found} 个可下载文件。")
        if progress_callback is not None:
            progress_callback(stage, processed_count, total_found, "")

        while in_flight:
            process_completed()
            if cancel_event is not None and cancel_event.is_set():
                for queued in in_flight:
                    queued.cancel()
                in_flight.clear()
                log(f"下载已取消，已处理 {processed_count}/{total_found} 个文件。")
                break

    return DownloadResult(
        total_found=total_found,
        downloaded_count=downloaded_count,
        skipped_existing_count=skipped_existing_count,
    )
