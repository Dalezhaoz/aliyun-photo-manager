import os
from dataclasses import dataclass


@dataclass
class OssConfig:
    cloud_type: str
    access_key_id: str
    access_key_secret: str
    endpoint: str
    bucket_name: str


def normalize_cloud_type(cloud_type: str) -> str:
    normalized = cloud_type.strip().lower()
    if normalized in {"aliyun", "oss", ""}:
        return "aliyun"
    if normalized in {"tencent", "cos", "tencent_cos"}:
        return "tencent"
    raise ValueError("暂不支持该云类型，请选择阿里云 OSS 或腾讯云 COS。")


def validate_oss_credentials(
    cloud_type: str,
    access_key_id: str,
    access_key_secret: str,
    endpoint: str,
) -> tuple[str, str, str, str]:
    normalized_cloud_type = normalize_cloud_type(cloud_type)
    missing = [
        name
        for name, value in (
            ("AccessKey ID", access_key_id.strip()),
            ("AccessKey Secret", access_key_secret.strip()),
            ("Endpoint / Region", endpoint.strip()),
        )
        if not value
    ]
    if missing:
        joined = ", ".join(missing)
        provider_name = "阿里云 OSS" if normalized_cloud_type == "aliyun" else "腾讯云 COS"
        raise ValueError(f"缺少 {provider_name} 配置：{joined}")
    return (
        normalized_cloud_type,
        access_key_id.strip(),
        access_key_secret.strip(),
        endpoint.strip(),
    )


def validate_oss_config(config: OssConfig) -> OssConfig:
    cloud_type, access_key_id, access_key_secret, endpoint = validate_oss_credentials(
        config.cloud_type,
        config.access_key_id,
        config.access_key_secret,
        config.endpoint,
    )
    bucket_name = config.bucket_name.strip()
    if not bucket_name:
        raise ValueError("请选择 bucket。")
    return OssConfig(
        cloud_type=cloud_type,
        access_key_id=access_key_id,
        access_key_secret=access_key_secret,
        endpoint=endpoint,
        bucket_name=bucket_name,
    )


def load_oss_config() -> OssConfig:
    return validate_oss_config(
        OssConfig(
            cloud_type=os.getenv("STORAGE_CLOUD_TYPE", "aliyun").strip(),
            access_key_id=os.getenv("OSS_ACCESS_KEY_ID", "").strip(),
            access_key_secret=os.getenv("OSS_ACCESS_KEY_SECRET", "").strip(),
            endpoint=os.getenv("OSS_ENDPOINT", "").strip(),
            bucket_name=os.getenv("OSS_BUCKET_NAME", "").strip(),
        )
    )
