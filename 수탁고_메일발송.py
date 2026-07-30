"""
FI운용본부 수탁고 현황 이메일 자동 발송 스크립트
==============================================
- 지정 경로에서 최신 수탁고 엑셀 파일을 탐색하여 첨부
- 수탁고 기준일과 연결되는 금리차트북 PDF를 탐색하여 함께 첨부
- 각 파일명의 날짜를 이메일 제목/본문에 자동 반영
- Outlook 서명을 Display() → HTMLBody 방식으로 자동 삽입
"""

import glob
import os
import re
from datetime import datetime, timedelta

import win32com.client


# ============================================================
# 1. 설정값 (수신자, 경로, 서명, 폰트 등)
#    - 수신자/참조자/폰트 변경 시 이 섹션만 수정하면 됩니다.
# ============================================================

# 수신자 (To)
TO_RECIPIENTS = "mylee@koreainvestment.com"

# 참조 (CC)
CC_RECIPIENTS = ";".join([
    "dongju.kim@koreainvestment.com",
    "ickhwan.cho@koreainvestment.com",
    "naraz@koreainvestment.com",
    "cheongho@koreainvestment.com",
    "mhyoon@koreainvestment.com",
])

# 첨부파일이 위치한 디렉토리
ATTACHMENT_DIR = r"K:\부서 공유\FI운용본부\본부 수탁고 현황"

# 첨부파일 패턴
AUM_FILENAME_PATTERN = r"FI운용본부수탁고_(\d{8})\.xlsx"
CHARTBOOK_FILENAME_PATTERN = r"금리차트북\((\d{4}\.\d{2}\.\d{2})\)\.pdf"

# ── 폰트 설정 ──
# 원본 메일 기준: 맑은 고딕 10pt
FONT_FAMILY = "'맑은 고딕', 'Malgun Gothic', sans-serif"
FONT_SIZE = "10pt"


# ============================================================
# 2. 첨부파일 탐색
# ============================================================

def find_latest_aum_file(directory: str, pattern: str) -> tuple[str, str]:
    """
    파일명에 포함된 yyyymmdd가 가장 큰 수탁고 엑셀 파일을 찾습니다.

    Returns:
        (파일 전체 경로, yyyymmdd 날짜 문자열)
    """
    search_path = os.path.join(directory, "FI운용본부수탁고_*.xlsx")
    files = glob.glob(search_path)

    if not files:
        raise FileNotFoundError(
            f"'{directory}' 경로에 'FI운용본부수탁고_*.xlsx' 파일이 없습니다."
        )

    dated_files = []
    regex = re.compile(pattern, re.IGNORECASE)

    for filepath in files:
        filename = os.path.basename(filepath)
        match = regex.fullmatch(filename)
        if not match:
            continue

        date_text = match.group(1)
        try:
            date_value = datetime.strptime(date_text, "%Y%m%d").date()
        except ValueError:
            # 예: 20260230처럼 실제로 존재하지 않는 날짜는 제외합니다.
            continue

        dated_files.append((filepath, date_text, date_value))

    if not dated_files:
        raise FileNotFoundError(
            f"'{directory}' 경로에 패턴과 날짜 형식이 유효한 수탁고 파일이 없습니다."
        )

    filepath, date_text, _ = max(dated_files, key=lambda item: item[2])
    return filepath, date_text


def find_chartbook_file(
    directory: str,
    aum_date_text: str,
) -> tuple[str, str, bool]:
    """
    수탁고 기준일과 연결되는 금리차트북 PDF를 찾습니다.

    수탁고 기준일 + 1일의 파일을 우선 선택합니다. 해당 파일이 없으면
    주말·휴일 또는 예외적인 파일명 날짜를 고려하여 수탁고 기준일과 같거나
    이후인 PDF 중 날짜가 가장 최신인 파일을 선택합니다.

    Returns:
        (PDF 전체 경로, yyyy.mm.dd 날짜 문자열, 기준일 + 1일 일치 여부)
    """
    search_path = os.path.join(directory, "금리차트북(*).pdf")
    files = glob.glob(search_path)

    if not files:
        raise FileNotFoundError(
            f"'{directory}' 경로에 '금리차트북(yyyy.mm.dd).pdf' 파일이 없습니다."
        )

    regex = re.compile(CHARTBOOK_FILENAME_PATTERN, re.IGNORECASE)
    candidates = []

    for filepath in files:
        filename = os.path.basename(filepath)
        match = regex.fullmatch(filename)
        if not match:
            continue

        date_text = match.group(1)
        try:
            date_value = datetime.strptime(date_text, "%Y.%m.%d").date()
        except ValueError:
            continue

        candidates.append((filepath, date_text, date_value))

    if not candidates:
        raise FileNotFoundError(
            f"'{directory}' 경로에 패턴과 날짜 형식이 유효한 금리차트북이 없습니다."
        )

    aum_date = datetime.strptime(aum_date_text, "%Y%m%d").date()
    expected_date = aum_date + timedelta(days=1)

    # 통상적인 날짜 관계(수탁고 기준일 + 1일)를 가장 먼저 찾습니다.
    for filepath, date_text, date_value in candidates:
        if date_value == expected_date:
            return filepath, date_text, True

    # 날짜가 같은 예외나 주말·휴일 간격을 허용하되,
    # 수탁고 기준일보다 오래된 차트북은 자동 첨부하지 않습니다.
    current_candidates = [
        item for item in candidates if item[2] >= aum_date
    ]
    if not current_candidates:
        raise FileNotFoundError(
            "수탁고 기준일과 같거나 이후인 금리차트북이 없습니다. "
            "오래된 금리차트북은 자동 첨부하지 않습니다."
        )

    filepath, date_text, _ = max(
        current_candidates,
        key=lambda item: item[2],
    )
    return filepath, date_text, False


# ============================================================
# 3. 본문 HTML 생성
# ============================================================

def build_body_html(
    aum_date_formatted: str,
    chartbook_date_formatted: str,
) -> str:
    """
    이메일 본문 HTML을 생성합니다 (서명 미포함).

    Args:
        aum_date_formatted:       'yyyy-mm-dd' 형식의 수탁고 기준일
        chartbook_date_formatted: 'yyyy.mm.dd' 형식의 금리차트북 날짜
    """
    style = f"font-family:{FONT_FAMILY}; font-size:{FONT_SIZE}; margin:0;"

    return (
        f'<p style="{style}">&nbsp;</p>'
        f'<p style="{style}">안녕하십니까<br>'
        f'FI운용본부 김현수입니다.</p>'
        f'<p style="{style}">{aum_date_formatted} 기준 본부 수탁고 및 '
        f'{chartbook_date_formatted} 금리차트북 보내드립니다.</p>'
        f'<p style="{style}">감사합니다.</p>'
        f'<p style="{style}">&nbsp;</p>'
    )


# ============================================================
# 4. 이메일 생성 및 발송 (Outlook COM)
# ============================================================

def create_and_send_email(
    subject: str,
    body_html: str,
    to: str,
    cc: str,
    attachment_paths: list[str],
    send: bool = False,
) -> None:
    """
    Outlook COM을 통해 이메일을 생성합니다.

    Args:
        subject:          이메일 제목
        body_html:        HTML 형식의 본문 (서명 미포함)
        to:               수신자 (세미콜론 구분)
        cc:               참조자 (세미콜론 구분)
        attachment_paths: 첨부파일 전체 경로 목록
        send:             True이면 즉시 발송, False이면 Display(미리보기)
    """
    missing_paths = [
        path for path in attachment_paths if not os.path.isfile(path)
    ]
    if missing_paths:
        raise FileNotFoundError(
            "다음 필수 첨부파일이 없습니다:\n" + "\n".join(missing_paths)
        )

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem

    mail.To = to
    mail.CC = cc
    mail.Subject = subject

    # Display()를 먼저 호출하여 Outlook 기본 서명을 삽입합니다.
    mail.Display()

    existing_html = mail.HTMLBody
    body_tag_match = re.search(r"(<body[^>]*>)", existing_html, re.IGNORECASE)

    if body_tag_match:
        insert_pos = body_tag_match.end()
        new_html = (
            existing_html[:insert_pos]
            + body_html
            + existing_html[insert_pos:]
        )
    else:
        new_html = body_html + existing_html

    mail.HTMLBody = new_html

    for attachment_path in attachment_paths:
        mail.Attachments.Add(attachment_path)

    if send:
        mail.Send()
        print("[완료] 이메일이 발송되었습니다.")
    else:
        print("[완료] 이메일이 Outlook에 표시되었습니다. 확인 후 발송해주세요.")


# ============================================================
# 5. 메인 실행부
# ============================================================

def main() -> None:
    # (1) 최신 수탁고 파일 탐색
    aum_filepath, aum_date_text = find_latest_aum_file(
        ATTACHMENT_DIR,
        AUM_FILENAME_PATTERN,
    )
    print(
        f"[정보] 수탁고 파일: {os.path.basename(aum_filepath)} "
        f"(기준일: {aum_date_text})"
    )

    # (2) 연결되는 금리차트북 탐색
    chartbook_filepath, chartbook_date_text, is_next_day = (
        find_chartbook_file(ATTACHMENT_DIR, aum_date_text)
    )
    print(
        f"[정보] 금리차트북: {os.path.basename(chartbook_filepath)} "
        f"(날짜: {chartbook_date_text})"
    )
    if not is_next_day:
        print(
            "[주의] 금리차트북 날짜가 수탁고 기준일의 다음 날이 아닙니다. "
            "Outlook 미리보기에서 첨부파일을 확인해주세요."
        )

    # (3) 날짜 포맷 변환: yyyymmdd → yyyy-mm-dd
    aum_date_formatted = datetime.strptime(
        aum_date_text,
        "%Y%m%d",
    ).strftime("%Y-%m-%d")

    # (4) 이메일 제목 및 본문 구성
    subject = f"[FI운용본부] FI본부수탁고_{aum_date_text} 기준"
    body_html = build_body_html(
        aum_date_formatted,
        chartbook_date_text,
    )

    # (5) 이메일 생성 (Display 모드)
    create_and_send_email(
        subject=subject,
        body_html=body_html,
        to=TO_RECIPIENTS,
        cc=CC_RECIPIENTS,
        attachment_paths=[aum_filepath, chartbook_filepath],
        send=False,  # True로 변경하면 즉시 발송
    )


if __name__ == "__main__":
    main()
