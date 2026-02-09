from PIL import Image
import os

def process_model_images(prdName, input_path, output_folder, zoom_factor=2.0, crop_specs=None):
    """
    input_path: 원본 이미지 경로
    output_folder: 저장될 폴더 경로
    zoom_factor: 확대 배율 (2.0 = 200%, 3.0 = 300%)
    crop_specs: 자를 영역 정보 리스트 [{'name': '부위명', 'x': x좌표, 'y': y좌표, 'w': 가로, 'h': 세로}]
    """
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 1. 이미지 불러오기
    img = Image.open(input_path)
    original_w, original_h = img.size

    # 2. 이미지 확대 (LANCZOS 필터로 화질 저하 최소화)
    new_size = (int(original_w * zoom_factor), int(original_h * zoom_factor))
    zoomed_img = img.resize(new_size, Image.Resampling.LANCZOS)

    # 3. 지정된 영역별로 크롭 및 저장
    for spec in crop_specs:
        name = spec['name']
        # 확대된 이미지 기준의 좌표 계산
        left = spec['x']
        top = spec['y']
        right = left + spec['w']
        bottom = top + spec['h']

        # 영역 자르기
        cropped_img = zoomed_img.crop((left, top, right, bottom))
        
        # 저장
        save_path = os.path.join(output_folder, f"{name}.jpg")
        cropped_img.save(save_path, quality=95)
        print(f"저장 완료: {save_path}")

prdName = [
  '뉴욕맨투맨',
  '댄디맨투맨',
  '러블리맨투맨',
  '모먼트맨투맨',
  '베이스맨투맨',
  '보스턴맨투맨',
  '브레이브맨투맨',
  '브루클린맨투맨',
  '슈페리어맨투맨',
  '시그니처맨투맨',
  '시트러스맨투맨',
  '올데이맨투맨',
  '웨스트맨투맨',
  '유나이티드맨투맨',
  '저니후드',
  '챔프맨투맨',
  '클래씨맨투맨',
  '테디맨투맨',
  '트루후드',
]

for i in prdName:
  
  try:
    
    crops = [
    {'name': '{}5'.format(i), 'x': 1060, 'y': 537, 'w': 1000, 'h': 500},  # 커버스티치
    ]
    
    crops2= [
    {'name': '{}6'.format(i), 'x': 520, 'y': 300, 'w': 500, 'h': 500}, # 상반신
    {'name': '{}8'.format(i), 'x': 350, 'y': 500, 'w': 500, 'h': 500}, # 왼쪽 소매
    {'name': '{}9'.format(i), 'x': 650, 'y': 500, 'w': 500, 'h': 500}, # 오른쪽 소매
    ]
    
    crops3= [
    {'name': '{}7'.format(i), 'x': 800, 'y': 500, 'w': 500, 'h': 500}, # 나염확대
    ]

    process_model_images(i,'./source/{}.jpg'.format(i), 'output_result', zoom_factor=3.0, crop_specs=crops)
    process_model_images(i,'./source/{}.jpg'.format(i), 'output_result', zoom_factor=1.5, crop_specs=crops2)
    process_model_images(i,'./source/{}.jpg'.format(i), 'output_result', zoom_factor=2.0, crop_specs=crops3)
    
  except Exception as e:
    print(e)