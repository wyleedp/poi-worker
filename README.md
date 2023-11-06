# poi-worker
Apache POI 라이브러리를 이용하여 MS Word(.docx) 및 Excel(.xlsx)을 생성하는 예제 프로젝트

### 개발환경
* JDK 11/Maven
* Apache POI 5.2.4

### 예제별 설명
1. com.github.wyleedp.poi.worker.WordUserHomeDirListCreate
    * 사용자 홈디텍토리의 폴더명 목록을 워드파일로 생성하는 예제
    * 생성된 워드파일은 사용자 임시폴더의 년월일시분초_UserHome.docx 파일로 생성된다.
        * 워드파일 경로 예) C:\Users\wyleedp\AppData\Local\Temp\20231106095505_UserHome.docx
2. com.github.wyleedp.poi.worker.WordHelloWorld
    * HelloWorld 문자열을 워드파일로 생성하는 예제
    * 생성된 워드파일은 사용자 임시폴더의 년월일시분초_HelloWorld.docx 파일로 생성된다.
        * 워드파일 경로 예) C:\Users\wyleedp\AppData\Local\Temp\20231106101051_HelloWorld.docx