import cv2
from pyzbar.pyzbar import decode
import openpyxl
import numpy as np

spreadSheetName = "Crescendo 2024 Scouting Data.xlsx"

def show_instruction_window():
    # Create a black image
    instruction_image = np.zeros((100, 700, 3), dtype=np.uint8)

    # Add text to the image
    instruction_text = "Press 'n' to scan another QR code or 'q' to finish scanning."
    cv2.putText(instruction_image, instruction_text, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1, cv2.LINE_AA)

    # Display the image
    cv2.imshow('Instructions', instruction_image)

def capture_and_decode_qr_code():
    cap = cv2.VideoCapture(0)  # 0 represents the default camera

    # Try to open the existing workbook
    try:
        workbook = openpyxl.load_workbook(spreadSheetName)
        sheet = workbook.active
    except FileNotFoundError:
        # Create a new workbook if the existing one is not found
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        # Add headers to the Excel sheet for a new workbook
        headers = ["EventKey", "MatchLevel", "MatchNumber", "Team", "ScouterName", "StartingPosition", "LeftStartingZone",
                   "Auto-AmpNotes", "Auto-SpeakerNotes", "Ground", "Source", "SpeakerNotes", "AmplifiedSpeakerNotes",
                   "AmpNotes", "Co-opBonus", "DefenseRating", "Tipped", "EndLocation", "FailedClimb", "RobotsOnstage",
                   "NoteScoredinTrap", "MechanicalFailures", "Comments", "ScoutedTime"]
        sheet.append(headers)

    qr_code_detected = False  # Flag to indicate if a QR code has been detected

    while True:
        try:
            while not qr_code_detected:
                ret, frame = cap.read()

                if ret:
                    decoded_objects = decode(frame)
                    for obj in decoded_objects:
                        data = obj.data.decode('utf-8')
                        process_qr_data(data, sheet)
                        print(f"QR Code Scanned: {data}")
                        qr_code_detected = True  # Set the flag to exit the loop

                    # Display on-screen prompt
                    cv2.putText(frame, 'Press "q" to finish scanning or any other key to scan another QR code',
                                (10, 20), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (255, 255, 255), 1, cv2.LINE_AA)
                    cv2.imshow('QR Code Scanner', frame)

                    key = cv2.waitKey(1)
                    if key == ord('q'):
                        break

            show_instruction_window()

            response = cv2.waitKey(0)
            if response == ord('n'):
                qr_code_detected = False  # Reset the flag for the next iteration
            elif response == ord('q'):
                break  # Exit the outer loop if the user wants to finish scanning

        except Exception as e:
            print(f"An error occurred: {e}")

    cap.release()
    cv2.destroyAllWindows()

    # Save the Excel workbook with a timestamped filename
    workbook.save(spreadSheetName)
    print(f"Data saved to {spreadSheetName}")

def process_qr_data(qr_data, sheet):
    # Split the QR data into individual fields
    fields = qr_data.split(',')

    # Specific data cleanup
    fields = fields[24:]
    fields.pop(3)
    fields.pop(4)
    fields[5] = fields[5] + fields.pop(6)
    temp = fields[1]
    fields[1] = fields[2]
    fields[2] = temp

    # Add QR data to the Excel sheet
    sheet.append(fields)

if __name__ == "__main__":
    capture_and_decode_qr_code()