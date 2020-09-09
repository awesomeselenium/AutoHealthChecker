# AutoHealthChecker
전국 시도교육청 통합 자가진단 hcs.eduro.go.kr 대응 자가진단 매크로  
절대 악용하지 말 것이며, 구두로 모든 학생의 상태를 물어본 후 자원이 부족할 경우에 사용할 것

## 엑셀 파일 작성 
끊기지 않게 [이름 생일 (한 열 비우고) 비밀번호] 형식으로 작성 
시/도, 학교급, 학교, 비밀번호는 각 행에 데이터가 없으면 1행을 기준으로 실행됨

|A|B|C|D|E|F|G|
|----|----|----|----|----|----|----|
|이름|생일|(실행 여부를 체크하는 칸)|시/도|학교급|학교|비밀번호|
||||||||

## 요구 사항
Selenium과 Openpyxl (pip install ~)으로 설치  
[Chrome webdriver](https://chromedriver.chromium.org/downloads) 다운받아서 .py 파일과 같은 폴더에 위치시킴  
studentlist.xlsx 파일도 같은 폴더에 위치  
[파이썬 최신 버전](https://www.python.org/downloads/)

## 개발해야할 부분 
- [ ] 컴퓨터의 Python과 설치되어있는 모듈에 의존하지 않는 EXE 형식으로 배포
- [ ] Typescript와 같은 언어로 크롬 띄우지 않고 자가진단 할 수 있게 구현
- [ ] 코드 가독성 수정

