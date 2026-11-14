"""本機續掃記錄：僅保存識別及內容雜湊，不保存帳密、Cookie 或表單內容。"""
import hashlib
import json
import os
from pathlib import Path
import tempfile


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False, default=str).encode()).hexdigest()


class ProgressStore:
    def __init__(self, path, temporary=False):
        self.path = Path(path)
        self.base_path = self.path
        self.temporary = temporary
        self.data = self.read(self.path) if not temporary else {'version':1,'jobs':{}}

    @staticmethod
    def read(path):
        if not path.exists():return {'version':1,'jobs':{}}
        try:
            data=json.loads(path.read_text(encoding='utf-8'))
            if data.get('version')!=1 or not isinstance(data.get('jobs'),dict):raise ValueError()
            if any(not isinstance(job,dict) or not isinstance(job.get('records'),dict) or not isinstance(job.get('rows'),dict) for job in data['jobs'].values()):raise ValueError()
            return data
        except (ValueError,OSError) as exc:
            raise ValueError('續掃暫存無法讀取；請備份 '+str(path)+' 後再處理，不會自動刪除或重設') from exc

    @classmethod
    def migrate_legacy(cls, legacy_path, directory):
        legacy_path=Path(legacy_path);directory=Path(directory)
        if not legacy_path.exists():return None
        data=cls.read(legacy_path)
        for key,job in data['jobs'].items():
            # 檔名只使用雜湊，不採用舊 JSON 中可能出現的任意路徑。
            if not isinstance(key,str) or len(key)!=64 or any(c not in '0123456789abcdef' for c in key):
                raise ValueError('舊進度的作業識別無法辨認，請先備份；未刪除舊檔')
            target=directory/('scan-'+key+'.json')
            if target.exists():cls.read(target)
            else:
                store=cls(target)
                store.data={'version':1,'jobs':{key:job}}
                store.flush()
        backup=legacy_path.with_name('progress.before-temp.json')
        count=1
        while backup.exists():
            count+=1;backup=legacy_path.with_name(f'progress.before-temp-{count}.json')
        legacy_path.rename(backup)
        return backup

    def begin(self, identity, rules, job_key=None):
        key = job_key or digest(identity)
        if job_key is not None:
            import re
            if not re.fullmatch(r'[0-9a-f]{64}',job_key):raise ValueError('案件進度識別格式不正確')
        self.job_key=key
        if self.temporary:
            self.path=self.base_path.parent/('scan-'+key+'.json')
            self.data=self.read(self.path)
        self.job = self.data['jobs'].setdefault(key, {'records': {}, 'rows': {}})
        stamp = digest(rules)
        if self.job.get('rules') != stamp:
            self.reset_completed(preserve_skipped=True)
            self.job['rules'] = stamp
            self.flush()
        if self.temporary:
            self.job['scan_state']='incomplete'
            self.job.setdefault('open_rows',{})
            self.flush()

    def reset_completed(self, preserve_skipped=False):
        # 重新掃描仍保留 pending 及待人工，避免未處理項目遺失或不明結果重送。
        self.job['rows'] = {}
        keep = ('pending', 'deferred', 'skipped') if preserve_skipped else ('pending', 'deferred')
        self.job['records'] = {k: v for k, v in self.job['records'].items() if v.get('status') in keep}
        self.flush()

    def retry_skipped(self):
        # 舊列快取沒有記錄每列包含哪些略過，重讀第一級，保留 done 與 pending。
        self.job['rows'] = {}
        self.job['records'] = {key: value for key, value in self.job['records'].items()
                               if value.get('status') != 'skipped'}
        self.flush()

    def defer(self, key, fingerprint, row_key, reason):
        old = self.record(key)
        entry = dict(old) if old.get('status') == 'pending' else {
            'status': 'deferred', 'fingerprint': fingerprint}
        entry.update(needs_manual=True, row_key=row_key, reason=reason)
        self.job['records'][key] = entry
        self.job['rows'].pop(row_key, None)
        self.flush()

    def manual_records(self):
        return {key: value for key, value in self.job['records'].items()
                if value.get('needs_manual') or value.get('status') == 'deferred'}

    def record(self, key):
        return self.job['records'].get(key, {})

    def done_row(self, key):
        return self.job['rows'].get(key) == 'done'

    def mark_row(self, key):
        self.job['rows'][key] = 'done'
        if self.temporary:self.job.setdefault('open_rows',{}).pop(key,None)
        self.flush()

    def mark(self, key, status, fingerprint, manual_row=None):
        entry = {'status': status, 'fingerprint': fingerprint}
        old = self.record(key)
        if status == 'pending' and old.get('needs_manual'):
            entry.update({name: old[name] for name in ('needs_manual', 'row_key', 'reason') if name in old})
        if status == 'pending' and manual_row is not None:
            entry.update(needs_manual=True, row_key=manual_row, reason='awaiting_verification')
            self.job['rows'].pop(manual_row, None)
        self.job['records'][key] = entry
        self.flush()

    def track_rows(self, rows, components=None):
        if not self.temporary:return
        old=self.job.get('open_rows',{})
        selected=set(components) if components is not None else None
        remaining={key:component for key,component in old.items() if selected is not None and component not in selected}
        remaining.update({key:component for key,component in rows.items() if not self.done_row(key)})
        self.job['open_rows']=remaining
        self.flush()

    def complete_scan(self):
        if not self.temporary:return False
        unresolved=any(r.get('status') in ('pending','deferred') or r.get('needs_manual') for r in self.job['records'].values())
        if self.job.get('open_rows') or unresolved:
            self.job['scan_state']='incomplete'
            self.flush()
            return False
        self.data['jobs'].pop(self.job_key,None)
        if self.data['jobs']:
            self.flush()
            return False
        self.path.unlink(missing_ok=True)
        return True

    def flush(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        fd, temporary = tempfile.mkstemp(prefix='progress-', suffix='.tmp', dir=self.path.parent)
        try:
            with os.fdopen(fd, 'w', encoding='utf-8') as stream:
                json.dump(self.data, stream, ensure_ascii=False)
                stream.flush()
                os.fsync(stream.fileno())
            os.replace(temporary, self.path)
        finally:
            if os.path.exists(temporary):
                os.unlink(temporary)
