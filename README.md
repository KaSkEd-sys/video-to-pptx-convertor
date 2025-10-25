# 🖼️ video to pptx convertor

Convert any **video** into a **PowerPoint presentation** — frame by frame!  
Perfect for creative experiments, meme slideshows, or visual effects inside PowerPoint.

---

## ✨ Features

- 🎥 Automatically extracts frames from any `.mp4` video  
- 🖼️ Saves each frame as an image in a folder  
- 🧩 Builds a PowerPoint presentation where **each slide = one frame**  
- 📊 Displays progress bars for extraction and slide creation  
- ⚡ Works offline — no internet required  
- 🪶 Lightweight and easy to use  

---

## 🧠 How it works

1. Loads your video file (`video.mp4`)  
2. Extracts every 10th frame (you can change this number)  
3. Creates a new `.pptx` file with all frames as slides  
4. Done! You get a full PowerPoint animation made from your video 🎬  

---

## 🛠️ Installation

```bash
pip install opencv-python python-pptx
```

## ✅ start

```bash
python badapple.py
```


## ⚙️ Optional Settings

Edit these lines in the script if needed:

```python
video_path = "video.mp4"               # Path to your video  
frames_folder = "base"                 # Folder for extracted frames  
output_ppt = "frames_to_slides.pptx"   # Output presentation name  
frame_step = 10                        # Save every 10th frame  
```
