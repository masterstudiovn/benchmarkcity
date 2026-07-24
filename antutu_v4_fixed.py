import re
import os
import time
import ast
from openpyxl import Workbook
from pyadb import adb
from xml.dom import minidom


# ---- Cấu hình ----
POLL_INTERVAL = 5          # giây, khoảng chờ mỗi lần kiểm tra benchmark đã xong chưa
MAX_WAIT = 420              # giây, thời gian chờ tối đa cho phần "chờ chấm điểm" (thay cho sleep(400) cứng)
OUTPUT_XLSX = 'antutu_result.xlsx'


# 从xml页面中提取分数
def get_score(path):
    score = []
    dom = minidom.parse(path)
    root = dom.documentElement
    nodes = root.getElementsByTagName('node')
    for node in nodes:
        score_node = node.getAttribute('resource-id')
        if score_node == 'com.antutu.ABenchMark:id/tv_score':
            text = node.getAttribute('text')
            score.append(text)
    return score


def _parse_bounds(bound_str):
    """Parse chuỗi bounds dạng '[x1,y1][x2,y2]' thành tâm điểm (px, py).
    Dùng ast.literal_eval thay vì eval() để an toàn hơn khi parse dữ liệu XML."""
    x, y = ast.literal_eval(re.sub(r"\]\[", "],[", bound_str))
    px = (x[0] + y[0]) / 2
    py = (x[1] + y[1]) / 2
    return px, py


# 从xml中提取点击坐标 —— trả về None, None nếu không tìm thấy, thay vì âm thầm giữ 0,0
def get_xml(path, pattern, by_resource_id=False):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f'Không tìm thấy file {path} — kiểm tra lại bước dump/pull trước đó.'
        )

    dom = minidom.parse(path)
    root = dom.documentElement
    nodes = root.getElementsByTagName('node')

    for node in nodes:
        if by_resource_id:
            attr_value = node.getAttribute('resource-id')
        else:
            attr_value = node.getAttribute('text')

        if attr_value == pattern:
            bound = node.getAttribute('bounds')
            px, py = _parse_bounds(bound)
            return px, py

    return None, None


def _dump_and_pull(remote_xml, local_xml, wait=2):
    """Gộp thao tác dump + pull, có kiểm tra tồn tại file cục bộ sau khi pull."""
    adb.shell(f'uiautomator dump /sdcard/{remote_xml}')
    time.sleep(wait)
    adb.pull(f'/sdcard/{remote_xml}', os.getcwd())
    local_path = os.path.join(os.getcwd(), local_xml)
    if not os.path.exists(local_path):
        raise FileNotFoundError(f'Pull thất bại: không thấy {local_path} sau khi dump {remote_xml}.')
    return local_path


def _tap(px, py):
    if px is None or py is None:
        raise ValueError('Không xác định được tọa độ tap (px/py là None) — pattern không khớp node nào trong XML.')
    adb.shell(f'input tap {px} {py}')


# 返回提取到的坐标
def operation(pattern):
    _dump_and_pull('2.xml', '2.xml')
    px, py = get_xml('2.xml', pattern)
    _tap(px, py)

    _dump_and_pull('3.xml', '3.xml')
    score = get_score('3.xml')
    px, py = get_xml('3.xml', pattern)
    _tap(px, py)

    return score


def _wait_for_total_score(remote_xml='1.xml', local_xml='1.xml'):
    """Poll thay cho sleep(400) cứng: kiểm tra định kỳ xem điểm Total đã xuất hiện chưa.
    Trả về (px, py, score) ngay khi tìm thấy, hoặc raise TimeoutError nếu quá MAX_WAIT."""
    waited = 0
    while waited < MAX_WAIT:
        try:
            local_path = _dump_and_pull(remote_xml, local_xml, wait=2)
        except FileNotFoundError:
            local_path = None

        if local_path:
            px, py = get_xml(local_xml, 'com.antutu.ABenchMark:id/tv_score', by_resource_id=True)
            if px is not None:
                score = get_score(local_xml)
                if score:
                    return px, py, score

        time.sleep(POLL_INTERVAL)
        waited += POLL_INTERVAL

    raise TimeoutError(
        f'Chờ quá {MAX_WAIT}s mà chưa thấy điểm Total — benchmark có thể đang treo hoặc UI đã đổi.'
    )


def score_operate():
    score_match = []

    adb.shell('am start com.antutu.ABenchMark/.ABenchMarkStart')
    time.sleep(10)

    _dump_and_pull('6.xml', '6.xml')
    oper_px, oper_py = get_xml('6.xml', '重新测试')
    _tap(oper_px, oper_py)

    # Chờ điểm Total bằng polling thay vì sleep(400) cứng.
    # Trước đây code còn đọc get_xml('1.xml', 0) — file 1.xml chưa từng
    # được tạo ra ở nhánh chạy thật (chỉ tồn tại nếu dòng shutil.copy được
    # bật lên), nên lỗi này sẽ FileNotFoundError ngay khi chạy. Giờ ta tự
    # dump/pull '1.xml' thật trong lúc poll, và tìm node bằng đúng
    # resource-id của điểm Total thay vì so sánh text với số 0.
    total_px, total_py, total_score = _wait_for_total_score()
    _tap(total_px, total_py)
    time.sleep(1)
    _tap(total_px, total_py)
    time.sleep(1)
    score_match += total_score

    score_match += operation('3D性能')
    time.sleep(2)

    score_match += operation('UX性能')
    time.sleep(2)

    score_match += operation('CPU性能')
    time.sleep(2)

    _dump_and_pull('8.xml', '8.xml')
    ram_score = get_score('8.xml')
    score_match += ram_score

    return score_match


def main():
    wb = Workbook()
    ws = wb.active
    ws.append(['Total',
               '3D',
               '3D[Graden]',
               '3D[Marooned]',
               'UX',
               'UX Data Secure',
               'UX Data Process',
               'UX Strategy games',
               'UX Image process',
               'UX I/O performance',
               'CPU',
               'CPU Mathematics',
               'CPU Common Use',
               'CPU Multi-Core',
               'RAM'])  # 添加所需数据名称

    try:
        score_match = score_operate()
    except Exception as e:
        # Trước đây: bất kỳ lỗi nào ở giữa cũng làm script chết mà không
        # lưu gì cả, dù đã tốn hàng chục phút chạy ADB. Giờ ta vẫn lưu file
        # Excel (chỉ có header) và in rõ lỗi để biết dừng ở bước nào.
        print(f'[LỖI] score_operate() dừng giữa chừng: {e}')
        wb.save(OUTPUT_XLSX)
        print(f'Đã lưu {OUTPUT_XLSX} (chỉ có header, do lỗi ở trên).')
        raise

    # Trước đây: score_match chỉ được print(), không có dòng nào ghi vào
    # ws và không có wb.save() — nên file Excel không bao giờ được tạo dù
    # Workbook đã setup header đầy đủ. Giờ ghi thẳng vào hàng dữ liệu và lưu.
    ws.append(score_match)
    wb.save(OUTPUT_XLSX)

    print(score_match)
    print(f'Đã lưu kết quả vào {OUTPUT_XLSX}')


if __name__ == '__main__':
    main()
