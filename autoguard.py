# -*- coding: utf-8 -*-

from datetime import datetime, timezone, timedelta
import time

import os

from xlwt import Workbook

import asyncio

import blivedm


class AutovipClient(blivedm.BLiveClient):
    def __init__(self, *args, **kargs):
        super().__init__(*args, **kargs)
        self.start_time = self._now().strftime('%Y_%m_%d_%H_%M_%S')
        self.last_checkpoint_filename = None
        self.workbook = Workbook()
        self.sheet = self.workbook.add_sheet('Membership')
        self.membership_count = 0
        self._add_to_sheet(['弹幕编号', 'B站ID', 'UID', '日期', '时间', '等级', '数量'])
        self.test_count = 0
        
    def _checkpoint(self):
        current_time = self._now().strftime('%Y_%m_%d-%H_%M_%S')
        filename = 'log/Membership-%s-%s.xls' % (self.start_time, current_time)
        self.workbook.save(filename)
        if self.last_checkpoint_filename is not None and os.path.isfile(self.last_checkpoint_filename):
            os.remove(self.last_checkpoint_filename)
        self.last_checkpoint_filename = filename

    def _now(self):
        return datetime.now(timezone(timedelta(hours=8))) # Set timezone to Beiing time

    def _level_to_text(self, level):
        level_map = {0: '非舰长', 1: '总督', 2: '提督', 3: '舰长'}
        return level_map[level]

    def _add_to_sheet(self, content: list):
        for index, item in enumerate(content):
            self.sheet.write(self.membership_count, index, item)
        self.membership_count += 1
    
    # def _create_test_record(self, message: blivedm.DanmakuMessage):
    #     self._add_to_sheet([
    #         self._now().strftime('%Y_%m_%d-%H_%M_%S') + '-' + str(message.uid), 
    #         self._now().strftime('%Y-%m-%d'), 
    #         self._now().strftime('%H:%M:%S'), 
    #         message.uid, 
    #         message.uname, 
    #         self._level_to_text(message.privilege_type), 
    #         1])
    #     self.test_count += 1
    #     if self.test_count % 10 == 0:
    #         self._checkpoint()
    #         self.test_count = 0

    def _create_record(self, message: blivedm.GuardBuyMessage):
        self._add_to_sheet([
            '%s-%s-%d-%d' % (message.uid, self._now().strftime('%Y_%m_%d_%H_%M_%S'), message.start_time, message.end_time),
            message.username, 
            message.uid, 
            self._now().strftime('%Y-%m-%d'), 
            self._now().strftime('%H:%M:%S'), 
            self._level_to_text(message.guard_level), 
            message.num])
        self._checkpoint()

    # async def _on_receive_danmaku(self, danmaku: blivedm.DanmakuMessage):
    #     print(f'{danmaku.uname}：{danmaku.msg}')
    #     self._create_test_record(danmaku)

    async def _on_buy_guard(self, message: blivedm.GuardBuyMessage):
        print(f'{message.username} 购买{message.gift_name}')
        self._create_record(message)


async def main():
    # 参数1是直播间ID
    # 如果SSL验证失败就把ssl设为False
    room_id = 22889484
    client = AutovipClient(room_id, ssl=True)
    future = client.start()
    try:
        # 5秒后停止，测试用
        # await asyncio.sleep(5)
        # future = client.stop()
        # 或者
        # future.cancel()

        await future
    finally:
        await client.close()


if __name__ == '__main__':
    asyncio.get_event_loop().run_until_complete(main())