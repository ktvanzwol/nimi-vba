Attribute VB_Name = "ni568x"
Option Explicit

Private Enum IviPwrMeter_AttributeIDs
    IVIPWRMETER_ATTR_CACHE = IVI_ATTR_CACHE
    IVIPWRMETER_ATTR_RANGE_CHECK = IVI_ATTR_RANGE_CHECK
    IVIPWRMETER_ATTR_QUERY_INSTRUMENT_STATUS = IVI_ATTR_QUERY_INSTRUMENT_STATUS
    IVIPWRMETER_ATTR_RECORD_COERCIONS = IVI_ATTR_RECORD_COERCIONS
    IVIPWRMETER_ATTR_SIMULATE = IVI_ATTR_SIMULATE
    IVIPWRMETER_ATTR_INTERCHANGE_CHECK = IVI_ATTR_INTERCHANGE_CHECK
    IVIPWRMETER_ATTR_SPY = IVI_ATTR_SPY
    IVIPWRMETER_ATTR_USE_SPECIFIC_SIMULATION = IVI_ATTR_USE_SPECIFIC_SIMULATION
    IVIPWRMETER_ATTR_CHANNEL_COUNT = IVI_ATTR_CHANNEL_COUNT
    IVIPWRMETER_ATTR_GROUP_CAPABILITIES = IVI_ATTR_GROUP_CAPABILITIES
    IVIPWRMETER_ATTR_FUNCTION_CAPABILITIES = IVI_ATTR_FUNCTION_CAPABILITIES
    IVIPWRMETER_ATTR_CLASS_DRIVER_PREFIX = IVI_ATTR_CLASS_DRIVER_PREFIX
    IVIPWRMETER_ATTR_CLASS_DRIVER_VENDOR = IVI_ATTR_CLASS_DRIVER_VENDOR
    IVIPWRMETER_ATTR_CLASS_DRIVER_DESCRIPTION = IVI_ATTR_CLASS_DRIVER_DESCRIPTION
    IVIPWRMETER_ATTR_CLASS_DRIVER_CLASS_SPEC_MAJOR_VERSION = IVI_ATTR_CLASS_DRIVER_CLASS_SPEC_MAJOR_VERSION
    IVIPWRMETER_ATTR_CLASS_DRIVER_CLASS_SPEC_MINOR_VERSION = IVI_ATTR_CLASS_DRIVER_CLASS_SPEC_MINOR_VERSION
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_PREFIX = IVI_ATTR_SPECIFIC_DRIVER_PREFIX
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_LOCATOR = IVI_ATTR_SPECIFIC_DRIVER_LOCATOR
    IVIPWRMETER_ATTR_IO_RESOURCE_DESCRIPTOR = IVI_ATTR_IO_RESOURCE_DESCRIPTOR
    IVIPWRMETER_ATTR_LOGICAL_NAME = IVI_ATTR_LOGICAL_NAME
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_VENDOR = IVI_ATTR_SPECIFIC_DRIVER_VENDOR
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_DESCRIPTION = IVI_ATTR_SPECIFIC_DRIVER_DESCRIPTION
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MAJOR_VERSION = IVI_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MAJOR_VERSION
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MINOR_VERSION = IVI_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MINOR_VERSION
    IVIPWRMETER_ATTR_INSTRUMENT_FIRMWARE_REVISION = IVI_ATTR_INSTRUMENT_FIRMWARE_REVISION
    IVIPWRMETER_ATTR_INSTRUMENT_MANUFACTURER = IVI_ATTR_INSTRUMENT_MANUFACTURER
    IVIPWRMETER_ATTR_INSTRUMENT_MODEL = IVI_ATTR_INSTRUMENT_MODEL
    IVIPWRMETER_ATTR_SUPPORTED_INSTRUMENT_MODELS = IVI_ATTR_SUPPORTED_INSTRUMENT_MODELS
    IVIPWRMETER_ATTR_CLASS_DRIVER_REVISION = IVI_ATTR_CLASS_DRIVER_REVISION
    IVIPWRMETER_ATTR_SPECIFIC_DRIVER_REVISION = IVI_ATTR_SPECIFIC_DRIVER_REVISION
    IVIPWRMETER_ATTR_DRIVER_SETUP = IVI_ATTR_DRIVER_SETUP
    IVIPWRMETER_ATTR_AVERAGING_AUTO_ENABLED = (IVI_CLASS_PUBLIC_ATTR_BASE + 3)
    IVIPWRMETER_ATTR_CORRECTION_FREQUENCY = (IVI_CLASS_PUBLIC_ATTR_BASE + 4)
    IVIPWRMETER_ATTR_OFFSET = (IVI_CLASS_PUBLIC_ATTR_BASE + 5)
    IVIPWRMETER_ATTR_RANGE_AUTO_ENABLED = (IVI_CLASS_PUBLIC_ATTR_BASE + 2)
    IVIPWRMETER_ATTR_UNITS = (IVI_CLASS_PUBLIC_ATTR_BASE + 1)
    IVIPWRMETER_ATTR_CHANNEL_ENABLED = (IVI_CLASS_PUBLIC_ATTR_BASE + 51)
    IVIPWRMETER_ATTR_RANGE_LOWER = (IVI_CLASS_PUBLIC_ATTR_BASE + 101)
    IVIPWRMETER_ATTR_RANGE_UPPER = (IVI_CLASS_PUBLIC_ATTR_BASE + 102)
    IVIPWRMETER_ATTR_TRIGGER_SOURCE = (IVI_CLASS_PUBLIC_ATTR_BASE + 201)
    IVIPWRMETER_ATTR_INTERNAL_TRIGGER_EVENT_SOURCE = (IVI_CLASS_PUBLIC_ATTR_BASE + 251)
    IVIPWRMETER_ATTR_INTERNAL_TRIGGER_LEVEL = (IVI_CLASS_PUBLIC_ATTR_BASE + 252)
    IVIPWRMETER_ATTR_INTERNAL_TRIGGER_SLOPE = (IVI_CLASS_PUBLIC_ATTR_BASE + 253)
    IVIPWRMETER_ATTR_AVERAGING_COUNT = (IVI_CLASS_PUBLIC_ATTR_BASE + 301)
    IVIPWRMETER_ATTR_DUTY_CYCLE_CORRECTION = (IVI_CLASS_PUBLIC_ATTR_BASE + 402)
    IVIPWRMETER_ATTR_DUTY_CYCLE_CORRECTION_ENABLED = (IVI_CLASS_PUBLIC_ATTR_BASE + 401)
    IVIPWRMETER_ATTR_REF_OSCILLATOR_ENABLED = (IVI_CLASS_PUBLIC_ATTR_BASE + 501)
    IVIPWRMETER_ATTR_REF_OSCILLATOR_FREQUENCY = (IVI_CLASS_PUBLIC_ATTR_BASE + 502)
    IVIPWRMETER_ATTR_REF_OSCILLATOR_LEVEL = (IVI_CLASS_PUBLIC_ATTR_BASE + 503)
End Enum

Public Enum ni568x_AttributeIDs
    NI568X_ATTR_RANGE_CHECK = IVI_ATTR_RANGE_CHECK
    NI568X_ATTR_QUERY_INSTRUMENT_STATUS = IVI_ATTR_QUERY_INSTRUMENT_STATUS
    NI568X_ATTR_CACHE = IVI_ATTR_CACHE
    NI568X_ATTR_SIMULATE = IVI_ATTR_SIMULATE
    NI568X_ATTR_RECORD_COERCIONS = IVI_ATTR_RECORD_COERCIONS
    NI568X_ATTR_INTERCHANGE_CHECK = IVI_ATTR_INTERCHANGE_CHECK
    NI568X_ATTR_SPECIFIC_DRIVER_PREFIX = IVI_ATTR_SPECIFIC_DRIVER_PREFIX
    NI568X_ATTR_SUPPORTED_INSTRUMENT_MODELS = IVI_ATTR_SUPPORTED_INSTRUMENT_MODELS
    NI568X_ATTR_GROUP_CAPABILITIES = IVI_ATTR_GROUP_CAPABILITIES
    NI568X_ATTR_INSTRUMENT_MANUFACTURER = IVI_ATTR_INSTRUMENT_MANUFACTURER
    NI568X_ATTR_INSTRUMENT_MODEL = IVI_ATTR_INSTRUMENT_MODEL
    NI568X_ATTR_INSTRUMENT_FIRMWARE_REVISION = IVI_ATTR_INSTRUMENT_FIRMWARE_REVISION
    NI568X_ATTR_SPECIFIC_DRIVER_REVISION = IVI_ATTR_SPECIFIC_DRIVER_REVISION
    NI568X_ATTR_SPECIFIC_DRIVER_VENDOR = IVI_ATTR_SPECIFIC_DRIVER_VENDOR
    NI568X_ATTR_SPECIFIC_DRIVER_DESCRIPTION = IVI_ATTR_SPECIFIC_DRIVER_DESCRIPTION
    NI568X_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MAJOR_VERSION = IVI_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MAJOR_VERSION
    NI568X_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MINOR_VERSION = IVI_ATTR_SPECIFIC_DRIVER_CLASS_SPEC_MINOR_VERSION
    NI568X_ATTR_LOGICAL_NAME = IVI_ATTR_LOGICAL_NAME
    NI568X_ATTR_IO_RESOURCE_DESCRIPTOR = IVI_ATTR_IO_RESOURCE_DESCRIPTOR
    NI568X_ATTR_DRIVER_SETUP = IVI_ATTR_DRIVER_SETUP
    NI568X_ATTR_CHANNEL_COUNT = IVI_ATTR_CHANNEL_COUNT
    NI568X_ATTR_UNITS = IVIPWRMETER_ATTR_UNITS
    NI568X_ATTR_RANGE_AUTO_ENABLED = IVIPWRMETER_ATTR_RANGE_AUTO_ENABLED
    NI568X_ATTR_AVERAGING_AUTO_ENABLED = IVIPWRMETER_ATTR_AVERAGING_AUTO_ENABLED
    NI568X_ATTR_CORRECTION_FREQUENCY = IVIPWRMETER_ATTR_CORRECTION_FREQUENCY
    NI568X_ATTR_OFFSET = IVIPWRMETER_ATTR_OFFSET
    NI568X_ATTR_AVERAGING_COUNT = IVIPWRMETER_ATTR_AVERAGING_COUNT
    NI568X_ATTR_RANGE_LOWER = IVIPWRMETER_ATTR_RANGE_LOWER
    NI568X_ATTR_RANGE_UPPER = IVIPWRMETER_ATTR_RANGE_UPPER
    NI568X_ATTR_DUTY_CYCLE_CORRECTION = IVIPWRMETER_ATTR_DUTY_CYCLE_CORRECTION
    NI568X_ATTR_DUTY_CYCLE_CORRECTION_ENABLED = IVIPWRMETER_ATTR_DUTY_CYCLE_CORRECTION_ENABLED
    NI568X_ATTR_TRIGGER_SOURCE = IVIPWRMETER_ATTR_TRIGGER_SOURCE
    NI568X_ATTR_INTERNAL_TRIGGER_EVENT_SOURCE = IVIPWRMETER_ATTR_INTERNAL_TRIGGER_EVENT_SOURCE
    NI568X_ATTR_INTERNAL_TRIGGER_LEVEL = IVIPWRMETER_ATTR_INTERNAL_TRIGGER_LEVEL
    NI568X_ATTR_INTERNAL_TRIGGER_SLOPE = IVIPWRMETER_ATTR_INTERNAL_TRIGGER_SLOPE
    NI568X_ATTR_INSTRUMENT_SERIAL_NUMBER = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 2)
    NI568X_ATTR_APERTURE_TIME_MODE = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 3)
    NI568X_ATTR_APERTURE_TIME = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 4)
    NI568X_ATTR_EXTERNAL_CALIBRATION_DATE = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 5)
    NI568X_ATTR_AVERAGING_AUTO_RESOLUTION = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 7)
    NI568X_ATTR_AVERAGING_AUTO_SOURCE = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 8)
    NI568X_ATTR_ENHANCED_MODULATION_MODE_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 9)
    NI568X_ATTR_BUFFER_SIZE = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 10)
    NI568X_ATTR_EXTERNAL_TRIGGER_EVENT_SOURCE = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 106)
    NI568X_ATTR_EXTERNAL_TRIGGER_SLOPE = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 107)
    NI568X_ATTR_TRIGGER_DELAY = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 120)
    NI568X_ATTR_TRIGGER_DELAY_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 121)
    NI568X_ATTR_TRIGGER_NOISE_IMMUNITY = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 122)
    NI568X_ATTR_TRIGGER_NOISE_IMMUNITY_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 123)
    NI568X_ATTR_TRIGGER_HYSTERESIS = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 124)
    NI568X_ATTR_TRIGGER_HYSTERESIS_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 125)
    NI568X_ATTR_TIME_SLOT_COUNT = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 200)
    NI568X_ATTR_TIME_SLOT_WIDTH = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 201)
    NI568X_ATTR_TIME_SLOT_EXCLUSION_START = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 202)
    NI568X_ATTR_TIME_SLOT_EXCLUSION_END = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 203)
    NI568X_ATTR_TIME_SLOT_EXCLUSION_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 204)
    NI568X_ATTR_SCOPE_RECORD_LENGTH = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 300)
    NI568X_ATTR_SCOPE_RECORD_POINTS = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 301)
    NI568X_ATTR_SCOPE_GATE_START = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 302)
    NI568X_ATTR_SCOPE_GATE_END = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 303)
    NI568X_ATTR_SCOPE_GATE_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 304)
    NI568X_ATTR_SCOPE_FENCE_START = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 305)
    NI568X_ATTR_SCOPE_FENCE_END = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 306)
    NI568X_ATTR_SCOPE_FENCE_ENABLED = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 307)
    NI568X_ATTR_SCOPE_GATE_AVERAGE_POWER = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 320)
    NI568X_ATTR_SCOPE_GATE_PEAK_POWER = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 321)
    NI568X_ATTR_SCOPE_GATE_MINIMUM_POWER = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 322)
    NI568X_ATTR_SCOPE_GATE_CREST_FACTOR = (IVI_SPECIFIC_PUBLIC_ATTR_BASE + 323)
End Enum

' Measurement Units
Public Enum ni568x_Units
    NI568X_VAL_DBM = 1         'Sets the units to dBm.
    NI568X_VAL_WATTS = 4       'Sets the units to watts.
    NI568X_VAL_MWATTS = 1001   'Sets the units to milliwatts.
    NI568X_VAL_UWATTS = 1002   'Sets the units to microwatts.
End Enum

Public Enum ni568x_ZeroStatus
    NI568X_VAL_ZERO_COMPLETE = 1         '/* Zero Correction Complete        */
    NI568X_VAL_ZERO_IN_PROGRESS = 0      '/* Zero Correction In Progress     */
    NI568X_VAL_ZERO_STATUS_UNKNOWN = -1  '/* Zero Correction Status Unknown  */
End Enum

' Time Limit Constants
Public Const NI568X_VAL_MAX_TIME_IMMEDIATE As Long = 0 'Immediate timeout.
Public Const NI568X_VAL_MAX_TIME_INFINITE As Long = -1 'Infinite timeout.

' ni568x Factory Method
Public Function ni568x_CreateSession(resourceName As String, Optional IDQuery As Boolean = True, Optional Reset As Boolean = True) As ni568x_Session
    Dim session As ni568x_Session
    
    Set session = New ni568x_Session
    session.InitSession resourceName, IDQuery, Reset
    
    Set ni568x_CreateSession = session
End Function


