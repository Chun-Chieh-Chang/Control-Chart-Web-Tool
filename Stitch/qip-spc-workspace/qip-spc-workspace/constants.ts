import { Anomaly, ChartDataPoint, KPI, Measurement, Status } from "./types";

export const MOCK_KPIS: KPI[] = [
  { label: 'Cpk', value: '1.33', subValue: 'Target: 1.67', status: 'neutral', icon: 'trend-up' },
  { label: 'Ppk', value: '1.28', subValue: 'Stable', status: 'good', icon: 'check-circle' },
  { label: 'Process Mean', value: '15.20', subValue: 'Target: 15.0', status: 'neutral', icon: 'minus' },
  { label: 'Standard Deviation', value: '0.85', subValue: 'Limit: 1.0', status: 'good', icon: 'activity' },
  { label: 'Anomaly Count', value: '15', subValue: 'Last 30 Days', status: 'bad', icon: 'alert-triangle' },
];

export const MOCK_ANOMALIES: Anomaly[] = [
  { id: '1', timestamp: '2023-05-17 12:59:55', message: 'Out of Control Limit - Sample 1180', sampleId: '1180', status: 'Pending' },
  { id: '2', timestamp: '2023-05-17 12:59:57', message: 'Out of Control Limit - Sample 1190', sampleId: '1190', status: 'Pending' },
  { id: '3', timestamp: '2023-05-17 12:59:58', message: 'Out of Control Limit - Sample 1180', sampleId: '1180', status: 'Pending' },
  { id: '4', timestamp: '2023-06-17 12:59:53', message: 'Out of Control Limit - Sample 1187', sampleId: '1187', status: 'Pending' },
  { id: '5', timestamp: '2023-06-17 12:59:51', message: 'Out of Control Limit - Sample 1180', sampleId: '1180', status: 'Pending' },
  { id: '6', timestamp: '2023-08-17 12:59:59', message: 'Out of Control Limit - Sample 1180', sampleId: '1180', status: 'Pending' },
];

export const MOCK_MEASUREMENTS: Measurement[] = [
  { id: '1', timestamp: '2023-05-11 18:39:35', sampleId: '1181', value: 39.5, operator: 'Operator', machine: 'Machine 1', status: Status.Normal },
  { id: '2', timestamp: '2023-05-17 18:39:38', sampleId: '1182', value: 25.6, operator: 'Operator', machine: 'Machine 1', status: Status.Normal },
  { id: '3', timestamp: '2023-05-11 18:33:36', sampleId: '1183', value: 78.7, operator: 'Operator', machine: 'Machine 1', status: Status.OutOfControl },
  { id: '4', timestamp: '2023-05-12 12:39:43', sampleId: '1184', value: 38.3, operator: 'Operator', machine: 'Machine 1', status: Status.Normal },
  { id: '5', timestamp: '2023-05-12 12:33:43', sampleId: '1185', value: 27.6, operator: 'Operator', machine: 'Machine 1', status: Status.Normal },
  { id: '6', timestamp: '2023-06-12 12:33:35', sampleId: '1186', value: 76.9, operator: 'Operator', machine: 'Machine 1', status: Status.OutOfControl },
  { id: '7', timestamp: '2023-06-12 12:33:35', sampleId: '1187', value: 79.5, operator: 'Operator', machine: 'Machine 2', status: Status.Normal },
  { id: '8', timestamp: '2023-06-12 12:33:35', sampleId: '1188', value: 79.5, operator: 'Operator', machine: 'Machine 2', status: Status.Normal },
  { id: '9', timestamp: '2023-06-12 12:33:45', sampleId: '1189', value: 25.7, operator: 'Operator', machine: 'Machine 2', status: Status.Normal },
];

// Generate chart data
const generateChartData = (): ChartDataPoint[] => {
  const data: ChartDataPoint[] = [];
  for (let i = 0; i < 30; i++) {
    const baseValue = 15.2;
    const noise = (Math.random() - 0.5) * 4;
    const spike = Math.random() > 0.9 ? (Math.random() > 0.5 ? 5 : -5) : 0;
    
    const val = baseValue + noise + spike;
    const range = Math.abs(noise * 2);

    data.push({
      name: `S${i + 1}`,
      value: parseFloat(val.toFixed(2)),
      ucl: 18.5,
      lcl: 12.5,
      mean: 15.2,
      range: parseFloat(range.toFixed(2)),
      rangeUcl: 12,
      rangeLcl: 0,
      rangeMean: 5
    });
  }
  return data;
};

export const MOCK_CHART_DATA = generateChartData();