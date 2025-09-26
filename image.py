import os
from huggingface_hub import InferenceClient
from dotenv import load_dotenv

load_dotenv()

client = InferenceClient(api_key=os.environ["HF_TOKEN"])

# output is a PIL.Image object
image = client.text_to_image(
    "Astronaut riding a horse",
    model="stabilityai/stable-diffusion-xl-base-1.0",
)

# Save the image locally
image.save("astronaut.png")
print("Image saved as astronaut.png")
