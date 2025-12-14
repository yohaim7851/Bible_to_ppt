import sqlite3
import os
import sys
import re # Import re module
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

def abbreviate_bible_book(book_name):
    """
    성경의 책 이름을 약어로 변환하는 함수.
    
    Args:
        book_name (str): 변환할 성경의 책 이름 (한글).

    Returns:
        str: 변환된 약어. 매칭되는 이름이 없으면 None을 반환.
    """
    
    abbreviations = {
        # 구약성경
        "창세기": "창",
        "출애굽기": "출",
        "레위기": "레",
        "민수기": "민",
        "신명기": "신",
        "여호수아": "수",
        "사사기": "삿",
        "룻기": "룻",
        "사무엘상": "삼상",
        "사무엘하": "삼하",
        "열왕기상": "왕상",
        "열왕기하": "왕하",
        "역대기상": "대상",
        "역대기하": "대하",
        "에스라": "스",
        "느헤미야": "느",
        "에스더": "에",
        "욥기": "욥",
        "시편": "시",
        "잠언": "잠",
        "전도서": "전",
        "아가": "아",
        "이사야": "사",
        "예레미야": "렘",
        "예레미야애가": "애",
        "에스겔": "겔",
        "다니엘": "단",
        "호세아": "호",
        "요엘": "욜",
        "아모스": "암",
        "오바댜": "옵",
        "요나": "욘",
        "미가": "미",
        "나훔": "나",
        "하박국": "합",
        "스바냐": "습",
        "학개": "학",
        "스가랴": "슥",
        "말라기": "말",
        
        # 신약성경
        "마태복음": "마",
        "마가복음": "막",
        "누가복음": "눅",
        "요한복음": "요",
        "사도행전": "행",
        "로마서": "롬",
        "고린도전서": "고전",
        "고린도후서": "고후",
        "갈라디아서": "갈",
        "에베소서": "엡",
        "빌립보서": "빌",
        "골로새서": "골",
        "데살로니가전서": "살전",
        "데살로니가후서": "살후",
        "디모데전서": "딤전",
        "디모데후서": "딤후",
        "디도서": "딛",
        "빌레몬서": "몬",
        "히브리서": "히",
        "야고보서": "약",
        "베드로전서": "벧전",
        "베드로후서": "벧후",
        "요한일서": "요일",
        "요한이서": "요이",
        "요한삼서": "요삼",
        "유다서": "유",
        "요한계시록": "계"
    }
    
    # 입력된 책 이름을 딕셔너리에서 찾아 약어로 변환
    book_name_trimmed = book_name.strip()
    if book_name_trimmed not in abbreviations:
        print(f"오류: '{book_name_trimmed}'는 등록되지 않은 성경 책 이름입니다.")
        return None
    
    return abbreviations.get(book_name_trimmed)

def generate_bible_ppt(query_string, output_filename="BiblePresentation.pptx", template_path=None):
    """
    성경 쿼리 문자열을 기반으로 PowerPoint 프레젠테이션을 생성합니다.

    Args:
        query_string (str): "창 1:1-5"와 같은 형식의 성경 쿼리 문자열.
        output_filename (str): 생성될 PowerPoint 파일의 이름.
        template_path (str, optional): 사용할 PowerPoint 템플릿 파일의 경로. 지정하지 않으면 기본 템플릿 사용.
    """
    print(f"--- Starting PPT Generation for query: {query_string} ---")

    # 스크립트가 실행되는 현재 디렉토리
    script_dir = Path(__file__).parent.resolve()

    # 데이터베이스 파일 경로 (스크립트 위치 기준)
    db_path = script_dir / 'Bibles' / 'bible.db'

    print(f"Database path: {db_path}")

    if not db_path.exists():
        print(f"FATAL ERROR: Database file not found at '{db_path}'")
        return

    # 템플릿 파일 경로 유효성 검사
    if template_path:
        template_full_path = Path(template_path).resolve()
        if not template_full_path.exists():
            print(f"FATAL ERROR: Template file not found at '{template_full_path}'.")
            return
        print(f"Using template: {template_full_path}")

    # 쿼리 파싱 (간단한 구현)
    # 예: "창 1:1-5" -> book_abbr="창", start_chapter=1, start_verse=1, end_verse=5
    parts = query_string.split()
    if len(parts) < 2:
        print("FATAL ERROR: Invalid query format. Expected 'BookAbbr Chapter:Verse-Verse' or 'BookAbbr Chapter:Verse'.")
        return

    book_abbr = abbreviate_bible_book(parts[0])
    chapter_verse_parts = parts[1].split(':')
    if len(chapter_verse_parts) < 1:
        print("FATAL ERROR: Invalid chapter/verse format.")
        return

    try:
        start_chapter = int(chapter_verse_parts[0])
        verse_range = chapter_verse_parts[1] if len(chapter_verse_parts) > 1 else "1" # Default to verse 1 if not specified

        if '-' in verse_range:
            start_verse, end_verse = map(int, verse_range.split('-'))
        else:
            start_verse = int(verse_range)
            end_verse = start_verse
    except ValueError:
        print("FATAL ERROR: Invalid chapter or verse number.")
        return

    # 데이터베이스 연결
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # 책 정보 가져오기
    cursor.execute("SELECT Id, Name, Abbreviation FROM Books WHERE Abbreviation = ? LIMIT 1", (book_abbr,))
    book_info = cursor.fetchone()

    if not book_info:
        print(f"FATAL ERROR: Book abbreviation '{book_abbr}' not found in database.")
        conn.close()
        return

    book_id, book_name, book_abbreviation = book_info

    # 구절 데이터 가져오기
    verses_data = []
    for chapter_num in range(start_chapter, start_chapter + 1): # 현재는 단일 장만 처리
        cursor.execute("""
            SELECT DISTINCT V.Number, V.Text, C.Number AS ChapterNumber
            FROM Verses V
            JOIN Chapters C ON V.ChapterId = C.Id
            JOIN Books B ON C.BookId = B.Id
            WHERE B.Id = ? AND C.Number = ? AND V.Number BETWEEN ? AND ?
            ORDER BY C.Number, V.Number
        """, (book_id, chapter_num, start_verse, end_verse))
        verses_data.extend(cursor.fetchall())

    conn.close()

    if not verses_data:
        print(f"No verses found for query: {query_string}")
        return

    # PowerPoint 프레젠테이션 생성
    try:
        if template_path:
            prs = Presentation(template_full_path)
        else:
            prs = Presentation()

        # 템플릿에 슬라이드 레이아웃이 있는지 확인
        if not prs.slide_layouts:
            print("FATAL ERROR: The template file contains no slide layouts. Please provide a valid template.")
            return

        # MyTemplate.pptx의 경우 'Title and Content' 레이아웃이 인덱스 0에 있습니다.
        slide_layout = prs.slide_layouts[0] 

        slide = prs.slides.add_slide(slide_layout)

        # 슬라이드 제목 텍스트 준비
        slide_title_text = ""
        if start_verse == end_verse:
            slide_title_text = f"({book_abbreviation} {start_chapter}:{start_verse})"
        else:
            slide_title_text = f"({book_abbreviation} {start_chapter}:{start_verse}-{end_verse})"

        # 슬라이드 본문 텍스트 준비 (모든 구절을 하나의 문자열로 합치기)
        combined_verse_lines = []
        for verse_num, verse_text, chapter_num in verses_data:
            combined_verse_lines.append(f"({book_abbreviation} {chapter_num}:{verse_num}) {verse_text}")
        combined_verse_content = "\n".join(combined_verse_lines)

        # 제목 플레이스홀더에 텍스트 할당
        # slide.shapes.title은 제목 플레이스홀더가 있을 경우에만 작동
        if slide.shapes.title:
            title = slide.shapes.title
            title.text = slide_title_text

        # 본문 플레이스홀더에 텍스트 할당
        # MyTemplate.pptx의 'Title and Content' 레이아웃에서 본문 플레이스홀더는 인덱스 1입니다.
        body_shape = slide.placeholders[1] 
        tf = body_shape.text_frame
        tf.text = combined_verse_content

        prs.save(output_filename)
        print(f"--- SUCCESS! Presentation saved to {output_filename} ---")

    except Exception as e:
        print("FATAL ERROR: An error occurred during PowerPoint generation.")
        print(f"Error details: {e}")
        print("Please ensure 'python-pptx' is installed and the template file is valid and contains appropriate layouts and placeholders.")




def get_user_input():
    """
    사용자로부터 성경 정보를 입력받는 함수.
    
    Returns:
        tuple: (query_string, output_filename, template_path)
    """
    print("=" * 50)
    print("성경 PPT 생성 프로그램")
    print("=" * 50)
    print()
    
    # 성경 책 이름 입력 (유효성 검사 포함)
    while True:
        book_name = input("성경 책 이름을 입력하세요 (예: 창세기, 마태복음): ").strip()
        if not book_name:
            print("책 이름을 입력해주세요.")
            continue
        
        # 책 이름이 abbreviations에 있는지 확인
        book_abbr = abbreviate_bible_book(book_name)
        if book_abbr is not None:
            break
        print("올바른 성경 책 이름을 입력해주세요. 다시 시도해주세요.")
        print()
    
    # 장 번호 입력
    while True:
        try:
            chapter = int(input("장 번호를 입력하세요 (예: 1): ").strip())
            if chapter > 0:
                break
            print("1 이상의 숫자를 입력해주세요.")
        except ValueError:
            print("올바른 숫자를 입력해주세요.")
    
    # 절 번호 입력
    while True:
        verse_input = input("절 번호를 입력하세요 (예: 15 또는 15-23): ").strip()
        if '-' in verse_input:
            try:
                start_verse, end_verse = map(int, verse_input.split('-'))
                if start_verse > 0 and end_verse >= start_verse:
                    verse_range = f"{start_verse}-{end_verse}"
                    break
                print("올바른 절 범위를 입력해주세요 (시작 절 <= 끝 절).")
            except ValueError:
                print("올바른 형식으로 입력해주세요 (예: 15-23).")
        else:
            try:
                verse = int(verse_input)
                if verse > 0:
                    verse_range = str(verse)
                    break
                print("1 이상의 숫자를 입력해주세요.")
            except ValueError:
                print("올바른 숫자를 입력해주세요.")
    
    # 쿼리 문자열 생성
    query_string = f"{book_name} {chapter}:{verse_range}"
    
    # 출력 파일명 입력
    output_filename = input("출력 파일명을 입력하세요 (예: bible.pptx, 기본값: bible.pptx): ").strip()
    if not output_filename:
        output_filename = "bible.pptx"
    if not output_filename.endswith('.pptx'):
        output_filename += '.pptx'
    
    # 템플릿 파일명 입력
    template_path = input("템플릿 파일명을 입력하세요 (예: MyTemplate.pptx, 기본값: MyTemplate.pptx, 없으면 Enter): ").strip()
    if not template_path:
        template_path = "MyTemplate.pptx"
    
    # 템플릿 파일 존재 여부 확인
    script_dir = Path(__file__).parent.resolve()
    template_full_path = script_dir / template_path
    if not template_full_path.exists():
        print(f"경고: 템플릿 파일 '{template_path}'을 찾을 수 없습니다. 기본 템플릿을 사용합니다.")
        template_path = None
    
    return query_string, output_filename, template_path


if __name__ == "__main__":
    # 스크립트가 있는 디렉토리를 기준으로 config.txt 파일의 경로를 설정합니다.
    script_dir = Path(__file__).parent.resolve()
    config_path = script_dir / 'config.txt'

    try:
        # 사용자 입력 받기
        query, output_name, template_arg = get_user_input()
        
        print()
        print("=" * 50)
        print(f"입력된 정보:")
        print(f"  성경 구절: {query}")
        print(f"  출력 파일: {output_name}")
        print(f"  템플릿 파일: {template_arg if template_arg else '기본 템플릿'}")
        print("=" * 50)
        print()
        
        # config.txt 파일 업데이트
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                f.write(f'"{query}"\n')
                f.write(f'"{output_name}"\n')
                f.write(f'"{template_arg if template_arg else ""}"\n')
            print(f"config.txt 파일이 업데이트되었습니다.")
        except Exception as e:
            print(f"경고: config.txt 파일 업데이트 중 오류가 발생했습니다: {e}")
        
        print()
        print("PPT 생성 중...")
        print()
        
        # 설정값으로 PPT 생성 함수를 호출합니다.
        generate_bible_ppt(query, output_name, template_arg)

    except KeyboardInterrupt:
        print("\n\n프로그램이 사용자에 의해 중단되었습니다.")
    except Exception as e:
        print(f"스크립트 실행 중 오류가 발생했습니다: {e}")
        import traceback
        traceback.print_exc()

