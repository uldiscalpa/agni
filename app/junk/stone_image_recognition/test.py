import cv2
import numpy as np


def generate_aruco_marker(aruco_dict, id, size, output_filename):
    # Generate the marker
    img = cv2.aruco.generateImageMarker(aruco_dict, id, size)

    # Save the marker to a file
    cv2.imwrite(output_filename, img)


def detect_and_annotate_markers(input_image_filename, output_image_filename):
    # Load the image
    image = cv2.imread(input_image_filename)

    if image is None:
        print(f"Failed to load image: {input_image_filename}")
        return

    # Initialize the detector with default parameters
    detector = cv2.aruco.ArucoDetector()

    # Use the default dictionary (you can change to others if needed)
    dictionary = cv2.aruco.getPredefinedDictionary(cv2.aruco.DICT_4X4_50)

    # Initialize empty lists for corners and rejectedImgPoints
    corners = []
    rejectedImgPoints = []
    ids = None  # For IDs, we'll use None as the initializer

    # Detect the markers
    corners, ids, rejectedImgPoints = detector.detectMarkers(
        image, corners, ids, rejectedImgPoints)

    # If markers are detected, annotate the image with the detected marker IDs
    if ids is not None and len(corners) > 0:
        cv2.aruco.drawDetectedMarkers(image, corners, ids)

    # Save the annotated image
    cv2.imwrite(output_image_filename, image)

    ##############################################
    image_for_test = 'test_tile_6.jpeg'
    image = cv2.imread(image_for_test)

    ##########
    # edge detection
    ############
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (3, 3), 0)
    # blurred = cv2.bilateralFilter(blurred, 9, 125, 255)
    thresh = cv2.adaptiveThreshold(
        blurred, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 3, 3)

    cv2.imwrite('preprocessed_image.jpg', thresh)
    contours, _ = cv2.findContours(
        thresh, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    min_contour_area = 1000  # Minimum contour area to keep
    filtered_contours = [
        c for c in contours if cv2.contourArea(c) > min_contour_area]
    print(len(filtered_contours))
    # Draw all detected contours on the original image
    # -1 in this context means draw all contours
    cv2.drawContours(image, filtered_contours, -1, (0, 255, 0), 2)

    # Save or display the image
    cv2.imwrite('contours_detected.jpg', image)


def test_creating_markers():

    # Assuming a resolution of 300 dpi which is ~118 pixels per cm
    marker_size_pixels = 10 * 118  # 10 cm

    # Create a custom ArUco dictionary object
    aruco_dict = cv2.aruco.getPredefinedDictionary(cv2.aruco.DICT_4X4_50)

    # Generate the ArUco markers
    generate_aruco_marker(
        aruco_dict, 0, marker_size_pixels, "aruco_marker_0.png")
    generate_aruco_marker(
        aruco_dict, 1, marker_size_pixels, "aruco_marker_1.png")

    print("ArUco markers generated and saved!")


def test_reading():
    input_image = "tile_test_image.jpg"
    output_image = "annotated_image.jpg"

    detect_and_annotate_markers(input_image, output_image)

    print(f"Annotated image saved as {output_image}!")


if __name__ == "__main__":
    test_creating_markers()
    test_reading()
