# feishu_doc.py
# -*- coding: utf-8 -*-
import logging
import json
import datetime
import requests
import lark_oapi as lark
from lark_oapi.api.docx.v1 import *
from typing import List, Dict, Any, Optional

# 读取配置（确保和你的config.py适配）
from src.config import config  # 若你的config是函数，替换为：from src.config import get_config; config = get_config()

# 初始化日志
logger = logging.getLogger(__name__)
logging.basicConfig(level=logging.INFO)

class FeishuDocManager:
    """飞书云文档管理器 + 机器人推送（整合版）"""
    def __init__(self):
        # 从配置文件读取飞书参数
        self.app_id = config.feishu_app_id
        self.app_secret = config.feishu_app_secret
        self.folder_token = config.feishu_folder_token
        self.bot_webhook = config.feishu_bot_webhook

        # 初始化飞书SDK客户端（自动处理token）
        if self.is_configured():
            self.client = lark.Client.builder() \
                .app_id(self.app_id) \
                .app_secret(self.app_secret) \
                .log_level(lark.LogLevel.INFO) \
                .build()
        else:
            self.client = None
            logger.warning("飞书配置不完整，SDK客户端未初始化")

    def is_configured(self) -> bool:
        """检查创建文档的核心配置是否完整"""
        return bool(self.app_id and self.app_secret and self.folder_token)

    def create_daily_doc(self, title: str, content_md: str) -> Optional[str]:
        """
        核心方法：创建飞书文档 + 写入Markdown内容 + 推送链接到飞书群
        :param title: 文档标题（如「2026-03-04 中线操盘复盘」）
        :param content_md: Markdown格式的文档内容
        :return: 文档链接（失败返回None）
        """
        # 1. 前置检查
        if not self.client or not self.is_configured():
            logger.error("飞书SDK未初始化，无法创建文档")
            return None

        try:
            # 2. 创建空文档
            create_request = CreateDocumentRequest.builder() \
                .request_body(CreateDocumentRequestBody.builder()
                              .folder_token(self.folder_token)
                              .title(title)
                              .build()) \
                .build()
            response = self.client.docx.v1.document.create(create_request)
            
            if not response.success():
                logger.error(f"创建空文档失败：{response.code} - {response.msg}")
                return None

            # 3. 获取文档ID和链接
            doc_id = response.data.document.document_id
            doc_url = f"https://feishu.cn/docx/{doc_id}"
            logger.info(f"空文档创建成功，链接：{doc_url}")

            # 4. 转换Markdown为飞书Block并写入
            blocks = self._markdown_to_sdk_blocks(content_md)
            self._batch_write_blocks(doc_id, blocks)
            logger.info("文档内容写入完成")

            # 5. 推送文档链接到飞书群（核心新增逻辑）
            if self.bot_webhook:
                self._send_doc_link_to_feishu(title, doc_url)
            else:
                logger.warning("飞书机器人Webhook未配置，跳过推送")

            return doc_url

        except Exception as e:
            logger.error(f"创建/推送文档异常：{str(e)}", exc_info=True)
            return None

    def _markdown_to_sdk_blocks(self, md_text: str) -> List[Block]:
        """Markdown转飞书SDK的Block对象（原有逻辑保留）"""
        blocks = []
        lines = md_text.split('\n')
        
        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 初始化默认值
            block_type = 2  # 2=普通文本
            text_content = line

            # 识别标题
            if line.startswith('# '):
                block_type = 3  # 3=H1
                text_content = line[2:]
            elif line.startswith('## '):
                block_type = 4  # 4=H2
                text_content = line[3:]
            elif line.startswith('### '):
                block_type = 5  # 5=H3
                text_content = line[4:]
            elif line.startswith('---'):
                # 分割线（22=Divider）
                blocks.append(Block.builder()
                              .block_type(22)
                              .divider(Divider.builder().build())
                              .build())
                continue

            # 构造文本元素
            text_run = TextRun.builder() \
                .content(text_content) \
                .text_element_style(TextElementStyle.builder().build()) \
                .build()
            
            text_element = TextElement.builder() \
                .text_run(text_run) \
                .build()
            
            text_obj = Text.builder() \
                .elements([text_element]) \
                .style(TextStyle.builder().build()) \
                .build()

            # 组装Block
            block_builder = Block.builder().block_type(block_type)
            if block_type == 2:
                block_builder.text(text_obj)
            elif block_type == 3:
                block_builder.heading1(text_obj)
            elif block_type == 4:
                block_builder.heading2(text_obj)
            elif block_type == 5:
                block_builder.heading3(text_obj)

            blocks.append(block_builder.build())

        return blocks

    def _batch_write_blocks(self, doc_id: str, blocks: List[Block]):
        """分批写入Block到文档（原有逻辑保留，优化命名）"""
        batch_size = 50  # 飞书API限制单次写入数量
        doc_block_id = doc_id  # 文档根节点ID就是文档ID
        
        for i in range(0, len(blocks), batch_size):
            batch_blocks = blocks[i:i+batch_size]
            # 构造写入请求
            add_request = CreateDocumentBlockChildrenRequest.builder() \
                .document_id(doc_id) \
                .block_id(doc_block_id) \
                .request_body(CreateDocumentBlockChildrenRequestBody.builder()
                              .children(batch_blocks)
                              .index(-1)  # -1=追加到末尾
                              .build()) \
                .build()
            
            resp = self.client.docx.v1.document_block_children.create(add_request)
            if not resp.success():
                logger.error(f"写入Block失败（批次{i//batch_size+1}）：{resp.code} - {resp.msg}")

    def _send_doc_link_to_feishu(self, title: str, doc_url: str):
        """
        核心新增：推送文档链接到飞书群（封装为私有方法）
        :param title: 文档标题
        :param doc_url: 文档链接
        """
        # 构造飞书机器人Markdown消息体
        msg_body = {
            "msg_type": "markdown",
            "content": {
                "title": "📋 操盘日报已生成",
                "text": f"""
### {title}
✅ 今日中线操盘复盘文档已创建完成，点击查看详情：
[📄 查看完整复盘文档]({doc_url})
---
> 生成时间：{datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
> 数据来源：TrendRadar 财经雷达
                """
            }
        }

        try:
            # 发送POST请求到飞书机器人
            response = requests.post(
                url=self.bot_webhook,
                headers={"Content-Type": "application/json"},
                data=json.dumps(msg_body),
                timeout=10  # 超时保护
            )
            response.raise_for_status()  # 抛出HTTP异常
            
            result = response.json()
            if result.get("code") == 0:
                logger.info("文档链接推送至飞书群成功")
            else:
                logger.error(f"推送失败：飞书返回错误 - {result}")

        except requests.exceptions.Timeout:
            logger.error("推送超时：飞书机器人服务未响应")
        except requests.exceptions.ConnectionError:
            logger.error("推送失败：无法连接到飞书机器人")
        except Exception as e:
            logger.error(f"推送异常：{str(e)}", exc_info=True)

# ------------------- 测试代码（可选，验证用） -------------------
if __name__ == "__main__":
    # 实例化管理器
    doc_manager = FeishuDocManager()
    
    # 测试用Markdown内容
    test_content = """
# 2026-03-04 中线操盘日报
## 1. 今日交易执行
- 持仓：光伏（20%）、半导体（15%）
- 无开仓/平仓操作，所有持仓均在20日均线上方

## 2. 市场分析
- 主流板块：光伏、储能（政策利好）
- 宏观：A50上涨0.5%，人民币汇率稳定

## 3. 明日计划
- 关注：储能板块回踩20日线的买点
- 风控：半导体若跌破20日线（18.5元），立即止损
---
### 核心原则
只做上升趋势，总仓位≤50%，单笔亏损≤1%
    """
    
    # 调用创建+推送方法
    doc_url = doc_manager.create_daily_doc(
        title="2026-03-04 中线操盘复盘",
        content_md=test_content
    )
    
    if doc_url:
        print(f"✅ 操作完成，文档链接：{doc_url}")
    else:
        print("❌ 创建/推送文档失败")
