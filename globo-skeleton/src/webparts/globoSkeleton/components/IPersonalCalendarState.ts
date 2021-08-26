import { IMeeting } from './IMeeting';

export interface IPersonalCalendarState {
    error: string;
    loading: boolean;
    meetings: IMeeting[];
}