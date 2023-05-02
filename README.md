# Klas_Macro
Only Access in KLAS DNS , so only use people who library workers and civil servants


사진을 누르면 유튜브 설명 영상으로 넘어갑니다.

[![Video Label](http://img.youtube.com/vi/LrTcQjTLAC4/0.jpg)](https://youtu.be/LrTcQjTLAC4)




<mark style='background-color: #ffdce0'> ** Chrome 이 깔려 있어야만 합니다.** </mark>

모든 준비를 완료하고 시작버튼을 누르면 자동화된 소프트웨어로 컨트롤 되는 Chrome 창이 뜨며 KLAS 홈페이지 에서 배가자료 관리 -> 파일 선택 -> Mark 수정 페이지로 이동하여, 서가정보 중 책 제목을 수정&저장 하는 프로그램입니다.

<mark style='background-color: #ffdce0'> 미리 Klas 사이트에서 배가자료관리의 자료 목록 보기 에서 1000개를 선택해 두세요 </mark>

<img width="80%" src="https://user-images.githubusercontent.com/61488231/235584736-f0e60840-8821-46e6-ae66-993db91de86f.png"/>

오직 KLAS 망에 있는 컴퓨터만 접속 할수 있습니다.

<mark style='background-color: #ffdce0'> 프로그램 제일 상단에 등록번호가 한줄씩 담긴 txt 파일을 추가하여 주십시오. </mark>

ex.) test.txt LA000000 LA000001 ....> LA000005

<mark style='background-color: #ffdce0'> KLAS 에 접속할 id 와 pw 를 입력하고 적용 버튼을 눌러 주십시오. </mark>

<img width="40%" src="https://user-images.githubusercontent.com/61488231/235585357-be93d2ba-d996-46e9-a933-7fc402e10d11.png"/>

** 보존번호를 부여하실떄 유용합니다. **

3-1. 첫번째 탭 에서는 엑셀파일을(.xlsx) 추가하여 주셔야 합니다. 해당 엑셀 파일에는 A열에서만 입력하실 텍스트를 인식합니다. 예를들어 책 제목 앞(혹은 뒤) 에 [보존번호] 를 붙히고 싶으시다면 [0001] 부터 [0100] 까지의 텍스트를 A1 부터 A100 까지의 엑셀의 셀에 저장하시고 [A 열만 해당됩니다.] 파일 탭에서 해당 엑셀파일을 추가 해 주십시오. 이후 제목의 앞에 올지 뒤에 올지를 체크박스를 통해 선택하시고 시작버튼을 누르시면 작동합니다.

<img width="80%" src="https://user-images.githubusercontent.com/61488231/235584881-36f45ef8-6351-4c2f-b38e-4a6a33b56c08.png"/>


책제목 : 티모와 함께하는 세계 여행 , 엑셀 A1 : [12345] 

앞 체크 = [12345] 티모와 함께하는 세계 여행 /

뒤 체크 = 티모와 함께하는 세계 여행 [12345]


** 출판사를 부여하실떄 유용합니다. **

4-1 두번째 탭 에서는 [ 텍스트 입력 ] 창에 임의로 부여하실 문자열을 입력하시고 적용을 눌러주십시오. 예를들어 [웅진] 을 입력하시고 저장하셨다고 하시면. 이후 제목의 앞(뒤) 에 "[웅진] 책제목" 으로 서사정보가 저장됩니다.

<img width="40%" src="https://user-images.githubusercontent.com/61488231/235585827-975e8f53-e1df-4c8f-87e9-ec431d37d79c.png"/>


책제목 : 티모와 함께하는 세계 여행 , 입력한 텍스트 : [웅진]

앞 체크 = [웅진] 티모와 함께하는 세계 여행 /

뒤 체크 = 티모와 함께하는 세계 여행 [웅진]


** 보존번호, 출판사 를 지우고 싶으실떄 유용합니다. **

4-2 세번째 탭 입니다. 격자괄호 [ ] 와 그 안에 존재하는 숫자 혹은 문자, 아니면 둘다 에 해당되는 문자를 제목에서 삭제합니다. 

예시:
  - 문자 체크 : [웅진] 
  - 숫자 체크 : [12345]
  - 둘다 체크 : [웅진] , [웅진12] , [12345]
  

<img width="40%" src="https://user-images.githubusercontent.com/61488231/235585943-c74f6c0d-859f-4a87-a79b-ecc212d3f13d.png"/>


등록번호.txt 파일에 입력된 책 목록 만큼 반복하여 웹 페이지에서 서지정보를 수정하게 됩니다.

주의사항 ::

KORMARC 245 필드 ▼a (서명) 부분만 수정 합니다. 서지 정보 오류시, 해당 페이지는 저장하지 않고 넘어갑니다.

이는 웹 페이지 안에서 동작하므로, 크롬창을 닫지 말아주십시오. ( 단, 다른 크롬창에서의 작업이나, 다른 프로그램 작업은 자유롭습니다 ) 닫기 버튼과 일시정지 버튼 작동이 미흡 할 수 있습니다. 예기치 못한 오류로 프로그램이 종료될 수 있습니다. ( 그러나 서지정보에 대한 잘못된 입력은 없습니다. ) Beta 버전으로 미완성된 프로그램 입니다. 추후 업데이트 예정입니다.

macro .zip 파일을 다운로드 받아 사용하세요
