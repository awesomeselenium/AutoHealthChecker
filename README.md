# AutoHealthChecker
전국 시도교육청 통합 자가진단 hcs.eduro.go.kr 대응 자가진단 매크로  

## 엑셀 파일 작성 
끊기지 않게 [이름 생일 (한 열 비우고) 비밀번호] 형식으로 작성  

## 요구 사항
Selenium과 Openpyxl (pip install ~)으로 설치  
[Chrome webdriver](https://chromedriver.chromium.org/downloads) 다운받아서 .py 파일과 같은 폴더에 위치시킴  
studentlist.xlsx 파일도 같은 폴더에 위치
[파이썬 최신 버전](https://www.python.org/downloads/)

## .py 파일에서 수정이 필요한 부분  
시,도 선택 부분에서 서울특별시 부터 "01", "02" ... 순으로 작성해야 함  
학교 급은 유치원부터 "1", "2",... 순으로 작성해야 함  
학교 이름은 띄어쓰기 없이 모두 작성해야 함 
