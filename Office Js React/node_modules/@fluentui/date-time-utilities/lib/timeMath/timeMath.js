import { TimeConstants } from '../dateValues/timeConstants';
/**
 * Returns a date offset from the given date by the specified number of minutes.
 * @param date - The origin date
 * @param minutes - The number of minutes to offset. 'minutes' can be negative.
 * @returns A new Date object offset from the origin date by the given number of minutes
 */
export var addMinutes = function (date, minutes) {
    var result = new Date(date.getTime());
    result.setTime(result.getTime() + minutes * TimeConstants.MinutesInOneHour * TimeConstants.MillisecondsIn1Sec);
    return result;
};
/**
 * Rounds the date's minute up to the next available increment. For example, if `date` has time 1:21
 * and `increments` is 5, the resulting time will be 1:25.
 * @param date - Date to ceil minutes
 * @param increments - Time increments
 * @returns Date with ceiled minute
 */
export var ceilMinuteToIncrement = function (date, increments) {
    var result = new Date(date.getTime());
    var minute = result.getMinutes();
    if (TimeConstants.MinutesInOneHour % increments) {
        result.setMinutes(0);
    }
    else {
        var times = TimeConstants.MinutesInOneHour / increments;
        for (var i = 1; i <= times; i++) {
            if (minute > increments * (i - 1) && minute <= increments * i) {
                minute = increments * i;
                break;
            }
        }
        result.setMinutes(minute);
    }
    return result;
};
//# sourceMappingURL=timeMath.js.map