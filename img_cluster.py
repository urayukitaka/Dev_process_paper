import numpy as np
import cv2

def labeling_cluster(img:np.array):

    # change data type
    img = img.astype("uint8")

    # img labeled 1
    img[img>=1] = 1

    # labeling in wide range
    # dilate
    kernel = np.ones((5,5),np.uint8)
    dilated = cv2.dilate(img, kernel, iterations = 1)

    # ラベルを付ける
    nlabel, img_lab = cv2.connectedComponents(dilated)

    # labeling
    labeled_img = img*img_lab
    labeled_img = labeled_img.astype("uint8")

    return nlabel, labeled_img