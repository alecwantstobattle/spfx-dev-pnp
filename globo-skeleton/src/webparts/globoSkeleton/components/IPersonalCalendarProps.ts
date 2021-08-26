import { IGloboSkeletonWebPartProps } from '../GloboSkeletonWebPart';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IPersonalCalendarProps extends IGloboSkeletonWebPartProps {
    graphClient: MSGraphClient
}