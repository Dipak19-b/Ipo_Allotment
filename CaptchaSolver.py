from tensorflow.keras.models import load_model
import tensorflow as tf
import numpy as np



model_kfintech = load_model("Model_KFINTECH.h5")
model_bigshare = load_model("Model_BIGSHARE.h5")

def encode_single_sample(img_path, label, crop):
    char_to_num = {'0': 0, '1': 1, '2': 2, '3': 3, '4': 4, '5': 5, '6': 6, '7': 7, '8': 8, '9': 9,
               'a': 10, 'b': 11, 'c': 12, 'd': 13, 'e': 14, 'f': 15, 'g': 16, 'h': 17, 'i': 18,
               'j': 19, 'k': 20, 'l': 21, 'm': 22, 'n': 23, 'o': 24, 'p': 25, 'q': 26, 'r': 27,
               's': 28, 't': 29, 'u': 30, 'v': 31, 'w': 32, 'x': 33, 'y': 34, 'z': 35,
               'A': 36, 'B': 37, 'C': 38, 'D': 39, 'E': 40, 'F': 41, 'G': 42, 'H': 43, 'I': 44,
               'J': 45, 'K': 46, 'L': 47, 'M': 48, 'N': 49, 'O': 50, 'P': 51, 'Q': 52, 'R': 53,
               'S': 54, 'T': 55, 'U': 56, 'V': 57, 'W': 58, 'X': 59, 'Y': 60, 'Z': 61}
    # Read image file and returns a tensor with dtype=string
    img = tf.io.read_file(img_path)
    # Decode and convert to grayscale (this conversion does not cause any information lost and reduces the size of the tensor)
    # This decode function returns a tensor with dtype=uint8
    img = tf.io.decode_png(img, channels=1)
    # Scales and returns a tensor with dtype=float32
    img = tf.image.convert_image_dtype(img, tf.float32)
    # Crop and resize to the original size : 
    # top-left corner = offset_height, offset_width in image = 0, 25 
    # lower-right corner is at offset_height + target_height, offset_width + target_width = 50, 150


    if(crop==True):
        img = tf.image.crop_to_bounding_box(img, offset_height=0, offset_width=25, target_height=50, target_width=125)
        img = tf.image.resize(img,size=[50,200],method='bilinear', preserve_aspect_ratio=False,antialias=False, name=None)


    # Transpose the image because we want the time dimension to correspond to the width of the image.
    img = tf.transpose(img, perm=[1, 0, 2])
    # Converts the string label into an array with 5 integers. E.g. '6n6gg' is converted into [6,16,6,14,14]
    label = list(map(lambda p:char_to_num[p], label))
    return img.numpy(), label


def predict_kfintech():
    # file = request.files['image']
    # #Resize image from image_path of pixel 200x50
    # im = Image.open(BytesIO(file.read()))
    # im = im.resize((200,50))
    # im.save("captcha_kfintech.png")

    a,b = encode_single_sample("captcha.png","123456",False)

    X=[]
    X.append(a)
    X=np.array(X)

    # print(X.shape)

    y_pred = model_kfintech.predict(X)
    y_pred = np.argmax(y_pred,axis=2)
    num_to_char = {'-1': 'UKN', '0': '0', '1': '1', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6', '7': '7', '8': '8', '9': '9'}

    tr = str(list(map(lambda x:num_to_char[str(x)], y_pred[0])))
    # print(tr)

    tr = ''.join(map(lambda x: num_to_char[str(x)], y_pred[0]))
    # my_int = int(tr)
    # print(my_int)
    # print(tr)
    return tr

def predict_Bigshare():
    # file = request.files['image']
    # #Resize image from image_path of pixel 200x50
    # im = Image.open(BytesIO(file.read()))
    # im = im.resize((200,50))
    # im.save("captcha_kfintech.png")

    a,b = encode_single_sample("captcha.png","0123456789",False)

    X=[]
    X.append(a)
    X=np.array(X)

    # print(X.shape)

    y_pred = model_bigshare.predict(X)
    y_pred = np.argmax(y_pred,axis=2)
    num_to_char = {'-1': 'UKN', '0': '0', '1': '1', '2': '2', '3': '3', '4': '4', '5': '5', '6': '6', '7': '7', '8': '8', '9': '9'}

    tr = str(list(map(lambda x:num_to_char[str(x)], y_pred[0])))
    # print(tr)

    tr = ''.join(map(lambda x: num_to_char[str(x)], y_pred[0]))
    # my_int = int(tr)
    # print(my_int)
    # print(tr)
    return tr

# print(predict_Bigshare())