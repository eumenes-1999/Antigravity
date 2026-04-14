import jwt
import datetime
import time
import requests
import json
import logging

logger = logging.getLogger(__name__)

class LineWorksAPI:
    def __init__(self, client_id, client_secret, service_account, private_key):
        self.client_id = client_id
        self.client_secret = client_secret
        self.service_account = service_account
        self.private_key = private_key
        self.access_token = None
        self.token_expiry = 0

    def _get_access_token(self):
        """JWTを使ってアクセストークンを取得する (API v2.0)"""
        if self.access_token and time.time() < self.token_expiry - 60:
            return self.access_token

        now = int(time.time())
        payload = {
            "iss": self.client_id,
            "sub": self.service_account,
            "iat": now,
            "exp": now + 3600
        }
        
        headers = {
            "alg": "RS256",
            "typ": "JWT"
        }

        # JWT作成
        encoded_jwt = jwt.encode(payload, self.private_key, algorithm="RS256", headers=headers)

        # トークン取得
        url = "https://auth.worksmobile.com/oauth2/v2.0/token"
        data = {
            "assertion": encoded_jwt,
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "calendar,calendar.read"
        }
        
        response = requests.post(url, data=data)
        if response.status_code != 200:
            logger.error(f"Failed to get access token: {response.text}")
            raise Exception(f"LineWorks Auth failed: {response.text}")

        res_data = response.json()
        self.access_token = res_data["access_token"]
        self.token_expiry = now + res_data["expires_in"]
        return self.access_token

    def get_candidate_events(self, user_id, keyword="候補", days_ahead=45):
        """指定したユーザーのカレンダーからキーワードを含むイベントを取得する"""
        token = self._get_access_token()
        
        now = datetime.datetime.utcnow()
        future = now + datetime.timedelta(days=days_ahead)
        
        # ISO8601形式 (UTC)
        start_time = now.strftime("%Y-%m-%dT%H:%M:%SZ")
        end_time = future.strftime("%Y-%m-%dT%H:%M:%SZ")
        
        # GASスクリプトで使用されていたURL形式
        url = f"https://www.worksapis.com/v1.0/users/{user_id}/calendar/events"
        params = {
            "rangeStartTime": start_time,
            "rangeEndTime": end_time,
            "limit": 100
        }
        
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(url, headers=headers, params=params)
        if response.status_code != 200:
            logger.error(f"Failed to fetch events: {response.text}")
            raise Exception(f"LineWorks API error: {response.text}")
            
        data = response.json()
        events = data.get("events", [])
        
        # キーワードでフィルタリング
        candidates = [e for e in events if keyword in (e.get("summary") or e.get("title") or "")]
        return candidates

    def format_events_as_list(self, events):
        """イベントリストを '・2026/04/03 12:00~17:00' 形式に整形する"""
        if not events:
            return ""
            
        formatted_lines = []
        
        # イベントを日付ごとにグルーピングするためのリストを作成
        parsed_events = []
        for ev in events:
            start = ev.get("start", {})
            end = ev.get("end", {})
            
            # dateTime or date
            start_str = start.get("dateTime") or start.get("date")
            end_str = end.get("dateTime") or end.get("date")
            
            if not start_str:
                continue
                
            # 文字列をdatetimeに変換 (UTC想定で、まずパース後にJST+9する)
            # LINE WORKSのAPIレスポンス形式に依存
            try:
                # 2026-04-03T12:00:00Z や 2026-04-03 の形式
                if 'T' in start_str:
                    dt_start = datetime.datetime.fromisoformat(start_str.replace('Z', '+00:00'))
                    dt_end = datetime.datetime.fromisoformat(end_str.replace('Z', '+00:00')) if end_str else None
                else:
                    dt_start = datetime.datetime.strptime(start_str, "%Y-%m-%d")
                    dt_end = datetime.datetime.strptime(end_str, "%Y-%m-%d") if end_str else None

                # 日本時間(JST)に変換 (簡易的に+9時間)
                jst_start = dt_start + datetime.timedelta(hours=9)
                jst_end = dt_end + datetime.timedelta(hours=9) if dt_end else None
                
                date_str = jst_start.strftime("%Y/%m/%d")
                time_range = ""
                if 'T' in start_str:
                    start_time = jst_start.strftime("%H:%M")
                    end_time = jst_end.strftime("%H:%M") if jst_end else ""
                    time_range = f" {start_time}~{end_time}"
                
                parsed_events.append({
                    "date": date_str,
                    "time": time_range,
                    "datetime": jst_start
                })
            except Exception as e:
                logger.warning(f"Error parsing event time: {e}")
                continue
        
        # 日付順にソート
        parsed_events.sort(key=lambda x: x["datetime"])
        
        for p in parsed_events:
            formatted_lines.append(f"・{p['date']}{p['time']}")
            
        return "\n".join(formatted_lines)
