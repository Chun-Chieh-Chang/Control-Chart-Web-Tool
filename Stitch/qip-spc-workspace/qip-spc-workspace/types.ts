export enum Status {
  Normal = 'Normal',
  OutOfControl = 'Out of Control',
  Warning = 'Warning'
}

export interface Measurement {
  id: string;
  timestamp: string;
  sampleId: string;
  value: number;
  operator: string;
  machine: string;
  status: Status;
}

export interface ChartDataPoint {
  name: string;
  value: number;
  ucl: number;
  lcl: number;
  mean: number;
  range?: number;
  rangeUcl?: number;
  rangeLcl?: number;
  rangeMean?: number;
}

export interface Anomaly {
  id: string;
  timestamp: string;
  message: string;
  sampleId: string;
  status: 'Pending' | 'Investigated' | 'Acknowledged';
}

export interface KPI {
  label: string;
  value: string;
  subValue?: string;
  status?: 'good' | 'bad' | 'neutral' | 'warning';
  icon?: string;
}