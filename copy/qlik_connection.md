# qlik_connection.py

```python
import ssl
import json
from websocket import create_connection
import logging

class QlikConnection:
    def __init__(self, settings):
        self.settings = settings
        self.ws = None
        self.doc_handle = None
        self.obj_handle = None

    def connect(self):
        """建立 WebSocket 连接"""
        self.ws = create_connection(
            self.settings.WS_URL,
            sslopt={"cert_reqs": ssl.CERT_NONE},
            header=self.settings.HEADERS
        )

        # 接收欢迎消息
        welcome_msg = self.ws.recv()
        return welcome_msg

    def request(self, method, params=None, handle=None, id=None):
        """发送 JSON-RPC 请求"""
        request = {
            "jsonrpc": "2.0",
            "method": method,
            "params": params or []
        }

        if handle is not None:
            request["handle"] = handle

        if id is not None:
            request["id"] = id

        self.ws.send(json.dumps(request))
        return self._wait_for_response(id)

    def _wait_for_response(self, expected_id):
        """异步等待指定ID的响应"""
        while True:
            raw_data = self.ws.recv()
            if not raw_data:
                continue

            response = json.loads(raw_data)

            # 检查是否是目标响应
            if response.get("id") == expected_id:
                return response

            # 处理推送来的通知消息（如Global类通知）
            if "method" in response:
                logging.info(f"收到推送通知: {response['method']}")
                continue

            logging.info(f"收到其他响应: ID={response.get('id')}")

    def close(self):
        """关闭连接"""
        if self.ws:
            self.ws.close()
```
