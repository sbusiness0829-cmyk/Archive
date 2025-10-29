# ComfyUI 실무 압축 가이드 (Windows · NVIDIA)

프로젝트에 바로 투입할 수 있도록 설치 → 핵심 워크플로 4종 → 확장 노드 → 속도/VRAM 최적화 → 실습 과제 → 트러블슈팅 순서로 정리했습니다. PowerShell 기준이며, 경로·명령은 상황에 맞게 조정하세요.

## 1) 환경 준비와 설치

- 권장 사양
	- GPU/VRAM: 8 GB 이상(SD 1.5), 12 GB+ 권장(SDXL)
	- NVIDIA 드라이버 최신(Studio 또는 Game Ready)
	- 디스크: 모델 보관용 20 GB+

- 설치 방법 A: Windows Portable(가장 빠름)
	1) ComfyUI Windows Portable 릴리스를 다운로드 후 압축 해제
	2) run_nvidia_gpu.bat 실행(최초 1회 의존성 자동 설치)

- 설치 방법 B: Git Clone(팀 표준/버전 고정에 유리)
	- 폴더 생성 후 클론, 의존성 설치, 실행

```powershell
# 선택 사항: Git 설치 후 실행
git clone https://github.com/comfyanonymous/ComfyUI.git
Set-Location ComfyUI
python -m pip install -r requirements.txt
python main.py
```

- 모델 폴더 구조(ComfyUI/models/)
	- checkpoints: 기반 모델(.safetensors/.ckpt)
	- vae: VAE 모델
	- loras: LoRA 가중치
	- controlnet: ControlNet/Annotator 관련
	- upscale_models: RealESRGAN/4x-UltraSharp 등
	- embeddings, clip, ipadapter 등 확장별 폴더

Tip: 첫날은 SD 1.5로 연습 → 둘째 날 SDXL 전환.

---

## 2) 핵심 워크플로 4종 익히기(노드 연결 순서)

아래는 최소 그래프입니다. 노드 이름은 기본 번들 기준이며, 상세 파라미터는 프로젝트에 맞게 조절합니다.

1) Text-to-Image (txt2img)
- Load Checkpoint → CLIP Text Encode(Positive/Negative) → KSampler → VAE Decode → Save Image
- 기본값 가이드: Steps 20–30, CFG 4–7, Sampler DPM++ 2M Karras, Seed 고정/랜덤

2) Image-to-Image (img2img)
- Load Checkpoint → VAE Encode(이미지) → KSampler(denoise 0.35–0.65) → VAE Decode → Save
- 해상도는 입력 이미지 유지, denoise로 스타일 전이 강도 조절

3) Upscale + Refiner
- Load Upscale Model(4x 등) → Upscale → Load Refiner Checkpoint → KSampler(Steps 8–12) → VAE Decode → Save
- 과도한 샤프/링잉 방지: Refiner Steps를 짧게, 샘플러는 Karras 계열 선호

4) ControlNet(정밀 가이드)
- Preprocessor(Canny/OpenPose/Depth 등) → ControlNet Loader → KSampler(Control 입력 연결) → VAE Decode → Save
- Control Weight 0.5±, Start/End 0.2–0.8 조절, 실패 시 Weight/CFG 낮추기

공통 파라미터 체크리스트: Resolution(512–768/SDXL=1024), Batch=1부터, Seed, Denoise Strength(img2img), CFG, Steps.

---

## 3) 확장 노드와 모델 세팅(우선순위 설치 목록)

- ComfyUI-Manager: 확장/업데이트 관리(필수)
- WAS Node Suite: 이미지 처리/유틸 다수
- Impact Pack: 고급 샘플러·유틸
- ControlNet + Aux(Annotator): Canny/Depth/OpenPose 전처리
- IPAdapter: 레퍼런스 이미지 스타일/ID 반영
- Tiled VAE / Tiled Diffusion: 고해상도 메모리 절약
- Upscale models: 4x-UltraSharp, RealESRGAN 계열

배치 경로(기본):
- ComfyUI/models/checkpoints, vae, loras, controlnet, upscale_models …

---

## 4) 속도·VRAM 최적화(현업에서 바로 쓰는 규칙)

- 해상도 우선 내리기(1024→768→640), Steps 16–24, Sampler DPM++ SDE/2M Karras
- Precision fp16, Tiled VAE/UNet 활성화(대형 해상도)
- SDXL Refiner는 필요 시에만(과샤프 방지), Batch=1 고정 후 Queue로 반복
- ControlNet 다중 사용 시: Weight 합산이 강하면 Steps 증가 대신 Weight↓
- OOM 발생 시 순서: 해상도↓ → ControlNet 수↓ → LoRA 수↓ → Refiner 생략

---

## 5) 실습 과제(총 90분) & 체크리스트

- A. txt2img(20분): 동일 프롬프트로 Seed 3개 생성 → 이미지/메타 저장
- B. img2img(15분): 기존 아트 1장 denoise 0.45 보정 → 전/후 비교 저장
- C. ControlNet Canny(20분): on/off 비교 → 에지 유지 확인 캡처
- D. Upscale+Refiner(20분): 768→1536 업스케일 → 과샤프 여부 확인
- E. 배치 큐(15분): 시드 5개 × 프롬프트 2개 큐잉 → 출력 폴더 정리

수락 기준(모든 과제 공통)
- 결과 이미지 1+ 장, 노드 그래프 캡처 1장, 워크플로 JSON/메타 저장(재현성)

---

## 6) 문제해결 가이드(자주 발생)

| 증상 | 가능 원인 | 해결 방법 |
|---|---|---|
| 모델을 못 찾음 | 모델 폴더 오배치/확장자 문제 | models/checkpoints 등 폴더 확인, .safetensors 권장, 파일명 특수문자 최소화 |
| 결과가 검게/노이즈만 | CFG 과다, Steps 부족, 샘플러 부적합 | CFG 4–7로 낮추기, Steps 24+, 샘플러 Karras 계열 교체 |
| CUDA OOM | 해상도 과다, ControlNet·LoRA 과다 | 해상도/ControlNet/LoRA 수 줄이기, Tiled VAE, fp16, SD 1.5로 전환 |
| 크기 불일치 에러 | 분기 해상도 불일치 | Encode→KSampler→Decode 경로 해상도 통일 |
| 지나친 샤프/경계 | Refiner 과다, 업스케일 모델 공격적 | Refiner Steps 8–12, 소프트 업스케일 모델 선택, 가우시안 소량 |
| 느리고 끊김 | 드라이버/확장 문제, 백그라운드 과다 | NVIDIA 드라이버 업데이트, 불필요 확장 제거, 재시동/클린 환경 |

---

### 부록: 실무 팁

- 이미지마다 “Save workflow/metadata” 저장 → 재현성 확보
- 프로젝트별 모델·LoRA·ControlNet 프리셋 폴더 관리
- 좋은 조합의 Torch/노드 버전은 폴더 스냅샷 백업
- 협업은 워크플로 JSON 공유 + 결과 폴더 규칙(날짜_프롬프트_시드)

