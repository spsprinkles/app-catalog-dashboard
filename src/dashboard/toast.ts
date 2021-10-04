/**
 * Toast
 */
export default class Toast {
    static showMsg(iconUrl: string, title: string, message: string) {
        var sip = null; //if user specific notification for presence info
        //First "" is to not show an extra span inline with title that has a smaller text size (odd use case)
        var notificationData = new window["SPStatusNotificationData"]("", window["STSHtmlEncode"](message), iconUrl, sip);
        var containerId = window["SPNotifications"].ContainerID.Status;
        var isSticky = false;
        var tooltip = null;
        var onClickCallback = null; //function
        var notification = new window["SPNotification"](containerId, window["STSHtmlEncode"](title), isSticky, tooltip, onClickCallback, notificationData);
        notification.Show(false); //false to animate show/hide
    }
}