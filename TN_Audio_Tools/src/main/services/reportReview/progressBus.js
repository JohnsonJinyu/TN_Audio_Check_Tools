const EventEmitter = require('events');
const bus = new EventEmitter();
bus.setMaxListeners(20);

const CHART_PROGRESS_EVENT = 'chart-progress';
const REVIEW_PROGRESS_EVENT = 'review-progress';

bus.events = {
	CHART_PROGRESS_EVENT,
	REVIEW_PROGRESS_EVENT,
};

bus.emitChartProgress = function(data) {
	bus.emit(CHART_PROGRESS_EVENT, data);
};

bus.emitReviewProgress = function(data) {
	bus.emit(REVIEW_PROGRESS_EVENT, data);
};

module.exports = bus;
