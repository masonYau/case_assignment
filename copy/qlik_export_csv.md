# qlik_export_csv.py

```python
import logging

import pandas as pd


class QlikExporter:
    def __init__(self, connection):
        self.conn = connection
        self.doc_handle = None  # 改由连接管理
        self._request_counter = 0

    def _next_id(self):
        self._request_counter += 1
        return self._request_counter

    def open_doc(self):
        """打开主文档"""
        resp = self.conn.request("OpenDoc", [self.conn.settings.APP_ID], -1, self._next_id())
        if 'result' in resp:
            self.doc_handle = resp['result']['qReturn']['qHandle']
            return True
        return False

    def init_report(self, bookmark_id):
        """初始化单个报告环境"""
        # 应用书签
        bookmark_resp = self.conn.request(
            "ApplyBookmark",
            [bookmark_id],
            self.doc_handle,
            self._next_id()
        )

        # 设置变量（两步）
        var_cache = {}
        for var_name in ["v_calc_Details", "v_Attributes"]:
            # 获取变量句柄
            var_resp = self.conn.request(
                "GetVariableByName",
                [var_name],
                self.doc_handle,
                self._next_id()
            )
            if "result" not in var_resp:
                continue

            var_handle = var_resp["result"]["qReturn"]["qHandle"]

            # 设置变量值
            self.conn.request(
                "SetNumValue",
                [-1 if var_name == "v_calc_Details" else 0],
                var_handle,
                self._next_id()
            )
        return True

    def _split_top_level(self, s: str):
        parts = []
        buf = []
        depth = 0
        in_squote = False
        in_dquote = False
        i = 0
        while i < len(s):
            ch = s[i]

            if ch == "'" and not in_dquote:
                in_squote = not in_squote
                buf.append(ch)
            elif ch == '"' and not in_squote:
                in_dquote = not in_dquote
                buf.append(ch)
            elif not in_squote and not in_dquote:
                if ch == "{":
                    depth += 1
                    buf.append(ch)
                elif ch == "}":
                    depth -= 1
                    buf.append(ch)
                elif ch == "," and depth == 0:
                    part = "".join(buf).strip()
                    if part:
                        parts.append(part)
                    buf = []
                else:
                    buf.append(ch)
            else:
                buf.append(ch)

            i += 1
        tail = "".join(buf).strip()
        if tail:
            parts.append(tail)
        return parts

    def _parse_set_modifier(self, bookmark_config_str: str):
        """
        输入形如：{<A={'x'},B={"=expr"},C={1,2}>}
        输出：
          {
            "A": {"type":"values","values":["x"]},
            "B": {"type":"search","search":"=expr"},
          }
        """
        s = bookmark_config_str.strip()

        # 去掉外层 {<  >}
        if s.startswith("{<") and s.endswith(">}"):
            s = s[2:-2].strip()
        elif s.startswith("<") and s.endswith(">"):
            s = s[1:-1].strip()

        items = self._split_top_level(s)

        result = {}
        for item in items:
            if "=" not in item:
                continue
            field, rhs = item.split("=", 1)
            field = field.strip()
            rhs = rhs.strip()

            # 期望 rhs 是 {...}
            if not (rhs.startswith("{") and rhs.endswith("}")):
                continue
            inner = rhs[1:-1].strip()

            # 去掉最外层一对引号（如 {"=xxx"} 或 {'a'}）
            # 注意：这里简化处理：只要是单元素且被引号包住就当 search 或 value
            if inner.startswith('"') and inner.endswith('"'):
                inner_unq = inner[1:-1]
                # if inner_unq.strip().startswith("="):
                #     result[field] = {"type": "search", "search": inner_unq}
                # else:
                #     result[field] = {"type": "values", "values": [inner_unq]}
                result[field] = {"type": "search", "search": inner_unq}
                continue
            # 多值：{'a','b'} 这种
            # 粗略按顶层逗号拆（内部没有再嵌套花括号时足够用）
            vals_raw = self._split_top_level(inner)
            values = []
            for v in vals_raw:
                v = v.strip()
                if (v.startswith("'") and v.endswith("'")) or (v.startswith('"') and v.endswith('"')):
                    v = v[1:-1]
                if v:
                    values.append(v)

            # 如果是唯一值且以=开头，也当高级搜索
            if len(values) == 1 and values[0].strip().startswith("="):
                result[field] = {"type": "search", "search": values[0]}
            else:
                result[field] = {"type": "values", "values": values}

        return result

    def _apply_selections(self, selections: dict):
        for field_name, rule in selections.items():

            # 取字段句柄
            f_resp = self.conn.request(
                "GetField",
                [field_name],
                self.doc_handle,
                self._next_id()
            )
            if "result" not in f_resp:
                continue
            f_handle = f_resp["result"]["qReturn"]["qHandle"]

            if rule["type"] == "search":
                # 高级搜索（你那种 "=Only({1}...)=Max(...)" / "=Field>=Min(...) and Field<=Max(...)"）
                search_str = rule["search"]

                # Field.SelectMatch 的参数在不同封装里可能略有差异；
                # 常见形态：SelectMatch([match], softlock)
                self.conn.request(
                    "Select",
                    [search_str, False, 0],
                    f_handle,
                    self._next_id()
                )

            else:
                values = rule.get("values", [])
                if not values:
                    continue
                q_field_values = [{"qText": v} for v in values]

                # 常见形态：SelectValues(qFieldValues, toggleMode, softLock)
                self.conn.request(
                    "SelectValues",
                    [q_field_values, False, False],
                    f_handle,
                    self._next_id()
                )

    def init_report_by_config(self, bookmark_config_str: str):
        """初始化单个报告环境（按配置字符串应用筛选，不用bookmark id）"""

        # 1) 清空选择
        self.conn.request("ClearAll", [], self.doc_handle, self._next_id())

        # 2) 应用配置选择
        selections = self._parse_set_modifier(bookmark_config_str)
        self._apply_selections(selections)

        # 3) 设置变量（保留你原逻辑）
        for var_name in ["v_calc_Details", "v_Attributes"]:
            var_resp = self.conn.request(
                "GetVariableByName",
                [var_name],
                self.doc_handle,
                self._next_id()
            )
            if "result" not in var_resp:
                continue
            var_handle = var_resp["result"]["qReturn"]["qHandle"]
            self.conn.request(
                "SetNumValue",
                [-1 if var_name == "v_calc_Details" else 0],
                var_handle,
                self._next_id()
            )

        return True

    def export_object(self, obj_id, output_path):
        """导出单个对象"""
        # 获取对象句柄
        obj_resp = self.conn.request(
            "GetObject",
            [obj_id],
            self.doc_handle,
            self._next_id()
        )
        if "result" not in obj_resp:
            return False
        obj_handle = obj_resp["result"]["qReturn"]["qHandle"]

        # 诊断对象
        layout = self.conn.request("GetLayout", [], obj_handle, self._next_id())
        layout_title = layout.get('result', dict()).get('qLayout', dict()).get('title')
        logging.info(f"Extracted Layout Title {layout_title}")

        if not self._validate_layout(layout):
            return False

        # 执行导出
        export_resp = self.conn.request(
            "ExportData",
            ["CSV_C", "/qHyperCubeDef", "temp.csv", "A"],
            obj_handle,
            self._next_id()
        )

        if "result" not in export_resp or "qUrl" not in export_resp["result"]:
            return False

        # 下载文件
        import requests
        url = f"https://{self.conn.settings.HOST}{export_resp['result']['qUrl']}"
        resp = requests.get(url, headers=self.conn.settings.HEADERS, verify=False)

        if resp.status_code == 200:
            with open(output_path, "wb") as f:
                f.write(resp.content)
            logging.info(f"Write {len(pd.read_csv(output_path))} records in {output_path}")
            return True
        return False

    def _validate_layout(self, layout):
        """验证对象布局"""
        if "result" not in layout:
            return False

        data = layout["result"]["qLayout"]
        obj_type = data["qInfo"]["qType"]

        if obj_type == "container":
            logging.info("错误: 检测到容器对象")
            return False

        if "qHyperCube" in data:
            rows = data["qHyperCube"]["qSize"]["qcy"]
            if rows == 0:
                logging.info("警告: 对象无数据")

        return True
```
