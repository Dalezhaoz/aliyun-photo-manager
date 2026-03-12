import hashlib
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Iterator, Optional, Set

try:
    from PIL import Image
    from PIL.ExifTags import TAGS
except ImportError:
    Image = None
    TAGS = {}


SUPPORTED_EXTENSIONS = {
    ".jpg",
    ".jpeg",
    ".png",
    ".heic",
    ".heif",
    ".bmp",
    ".gif",
    ".tif",
    ".tiff",
    ".webp",
}


@dataclass
class PhotoInfo:
    source: Path
    taken_time: datetime
    image_format: str
    orientation: str
    file_hash: str


def iter_photos(source_dir: Path) -> Iterator[Path]:
    for path in source_dir.rglob("*"):
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS:
            yield path


def get_exif_datetime(path: Path) -> Optional[datetime]:
    if Image is None:
        return None
    try:
        with Image.open(path) as image:
            exif = image.getexif()
            if not exif:
                return None
            tag_map = {TAGS.get(tag_id, tag_id): value for tag_id, value in exif.items()}
            for key in ("DateTimeOriginal", "DateTimeDigitized", "DateTime"):
                value = tag_map.get(key)
                if not value:
                    continue
                return datetime.strptime(str(value), "%Y:%m:%d %H:%M:%S")
    except Exception:
        return None
    return None


def get_orientation(path: Path) -> str:
    if Image is None:
        return "unknown"
    try:
        with Image.open(path) as image:
            width, height = image.size
    except Exception:
        return "unknown"
    if width > height:
        return "landscape"
    if width < height:
        return "portrait"
    return "square"


def compute_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_obj:
        while True:
            chunk = file_obj.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def build_photo_info(path: Path) -> PhotoInfo:
    taken_time = get_exif_datetime(path)
    if taken_time is None:
        taken_time = datetime.fromtimestamp(path.stat().st_mtime)
    return PhotoInfo(
        source=path,
        taken_time=taken_time,
        image_format=path.suffix.lower().lstrip(".") or "unknown",
        orientation=get_orientation(path),
        file_hash=compute_hash(path),
    )


def build_destination(photo: PhotoInfo, target_root: Path, flat: bool) -> Path:
    date_folder = photo.taken_time.strftime("%Y-%m-%d")
    if flat:
        base_dir = target_root / date_folder
    else:
        base_dir = (
            target_root
            / photo.taken_time.strftime("%Y")
            / photo.taken_time.strftime("%m")
            / date_folder
        )
    return base_dir / photo.image_format / photo.orientation


def ensure_unique_path(destination_dir: Path, source_name: str) -> Path:
    candidate = destination_dir / source_name
    if not candidate.exists():
        return candidate

    stem = Path(source_name).stem
    suffix = Path(source_name).suffix
    index = 1
    while True:
        candidate = destination_dir / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
        index += 1


def transfer_file(source: Path, destination: Path, move_files: bool, dry_run: bool) -> None:
    if dry_run:
        return
    destination.parent.mkdir(parents=True, exist_ok=True)
    if move_files:
        shutil.move(str(source), str(destination))
    else:
        shutil.copy2(source, destination)


def sort_photos(
    source_dir: Path,
    target_dir: Path,
    dry_run: bool = False,
    flat: bool = False,
    include_duplicates: bool = False,
    move_files: bool = False,
    logger: Optional[Callable[[str], None]] = None,
) -> None:
    seen_hashes: Set[str] = set()
    total = 0
    duplicate_count = 0
    def log(message: str) -> None:
        if logger is not None:
            logger(message)
        else:
            print(message)

    for photo_path in iter_photos(source_dir):
        total += 1
        photo = build_photo_info(photo_path)
        if not include_duplicates and photo.file_hash in seen_hashes:
            duplicate_count += 1
            log(f"[SKIP] 重复文件：{photo.source}")
            continue
        seen_hashes.add(photo.file_hash)

        destination_dir = build_destination(photo, target_dir, flat)
        destination_path = ensure_unique_path(destination_dir, photo.source.name)
        action = "MOVE" if move_files else "COPY"
        log(f"[{action}] {photo.source} -> {destination_path}")
        transfer_file(photo.source, destination_path, move_files, dry_run)

    log("")
    log(f"扫描照片：{total}")
    log(f"跳过重复：{duplicate_count}")
    log(f"输出目录：{target_dir}")
    if Image is None:
        log("提示：未安装 Pillow，当前使用文件修改时间，横竖图识别可能为 unknown。")
