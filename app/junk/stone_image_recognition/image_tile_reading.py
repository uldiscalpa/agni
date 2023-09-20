import cv2
import numpy as np

# Step 1: Generate ArUco markers
aruco_dict = cv2.aruco.Dictionary_get(cv2.aruco.DICT_6X6_250)
marker_size = 10  # Size in centimeters
marker_id_1 = 10
marker_id_2 = 20

# Generate marker images
marker_image_1 = cv2.aruco.drawMarker(aruco_dict, marker_id_1, 200)
marker_image_2 = cv2.aruco.drawMarker(aruco_dict, marker_id_2, 200)

# Step 2: Read the photo of the tile with markers
image_path = 'tile_with_markers.jpg'  # Replace with your image path
image = cv2.imread(image_path)

# Step 3: Flatten the image based on markers
corners_1, _ = cv2.aruco.detectMarkers(image, aruco_dict)
rvecs_1, tvecs_1, _ = cv2.aruco.estimatePoseSingleMarkers(corners_1, marker_size, camera_matrix, dist_coeffs)

corners_2, _ = cv2.aruco.detectMarkers(image, aruco_dict)
rvecs_2, tvecs_2, _ = cv2.aruco.estimatePoseSingleMarkers(corners_2, marker_size, camera_matrix, dist_coeffs)

# Perform perspective transformation based on tvecs
# You'll need to calculate the transformation matrix based on the marker positions

# Step 4: Calculate the area of the tile
# Apply image processing techniques to calculate the area while excluding cut-out parts

# Step 5: Save the image with contours for testing
image_with_contours = image.copy()
# Apply contour detection and draw contours on the image_with_contours

# Save the images
cv2.imwrite('marker_1.jpg', marker_image_1)
cv2.imwrite('marker_2.jpg', marker_image_2)
cv2.imwrite('flattened_tile.jpg', flattened_image)
cv2.imwrite('tile_with_contours.jpg', image_with_contours)