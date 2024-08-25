"""
invest_map.py 학생 거주 위치 지도 만들기   Ver 1.0_240518
"""
import openpyxl
import folium

# 엑셀 파일 로드
wb = openpyxl.load_workbook('./school/신상_자료.xlsx', data_only=True)  # 수식 결과 값을 읽기 위해 data_only=True 사용
ws = wb.active  # 첫 번째 시트를 활성화

# 데이터 읽기 (4번째 행부터 시작)
data = []
for row in ws.iter_rows(min_row=3, values_only=True):
    try:
        lat = row[41]  # 위도 값 가져오기
        lng = row[42]  # 경도 값 가져오기
        if lat is not None and lng is not None:
            lat = float(lat)  # 위도를 float으로 변환
            lng = float(lng)  # 경도를 float으로 변환
            data.append({
                '학번': row[0],
                '성명': row[1],
                '위도': lat,
                '경도': lng,
            })
        else:
            print(f"Missing coordinates for {row[1]}")
    except (TypeError, ValueError):
        # 변환 오류가 발생하면 해당 데이터를 건너뜁니다.
        print(f"Invalid data for {row[1]}: {row[2]}, {row[3]}")
        continue

# 수원 지도 만들기
station_map = folium.Map(location=[37.2711, 127.0469], zoom_start=16)

# 위치정보를 Marker로 표시
for entry in data:
    name = entry['성명']
    lat = entry['위도']
    lng = entry['경도']
    folium.Marker([lat, lng],
                  icon=folium.Icon(color="red", icon='info-sign'),
                  popup=name).add_to(station_map)

# 지도를 HTML 파일로 저장하기
station_map.save('./school/invest/신상_주소위치.html')
print("\nData saving....")


