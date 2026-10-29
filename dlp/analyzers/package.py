import io
import zipfile


def analyze_package(data, visible_bytes, text_bytes, value_bytes, policy):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        expanded = sum(e.file_size for e in entries)
        media = sum(e.file_size for e in entries if e.filename.startswith('xl/media/'))
        media_compressed = sum(e.compress_size for e in entries if e.filename.startswith('xl/media/'))
        embeddings = sum(e.file_size for e in entries if e.filename.startswith('xl/embeddings/'))
    ratio = (len(data) - media_compressed) / max(visible_bytes, 1)
    metrics = {'file_bytes': len(data), 'expanded_bytes': expanded, 'media_bytes': media,
               'embedded_object_bytes': embeddings,
               'text_utf8_bytes': text_bytes, 'serialized_value_utf8_bytes': value_bytes,
               'visible_value_utf8_bytes_estimate': visible_bytes,
               'non_media_file_to_visible_ratio': ratio}
    findings = []
    if len(data) - media_compressed >= policy.size_min_bytes and ratio >= policy.size_ratio:
        findings.append({'rule': 'file_content_size_ratio', 'family': 'size', 'score': 25})
    return metrics, findings
