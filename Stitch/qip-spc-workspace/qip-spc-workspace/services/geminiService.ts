import { GoogleGenAI } from "@google/genai";
import { Anomaly } from "../types";

const getClient = () => {
  const apiKey = process.env.API_KEY || ''; // In a real app, ensure this is set
  return new GoogleGenAI({ apiKey });
};

export const analyzeAnomalyWithAI = async (anomaly: Anomaly, contextData: any): Promise<string> => {
  try {
    const ai = getClient();
    if (!process.env.API_KEY) {
        return "Simulation: Gemini API Key not found. Please set REACT_APP_GEMINI_API_KEY (or process.env.API_KEY) to use real AI analysis. \n\n Simulated Analysis: The deviation in Sample " + anomaly.sampleId + " suggests a potential calibration drift in Sensor array B. Recommend immediate recalibration.";
    }

    const prompt = `
      You are an expert Manufacturing Quality Control Engineer.
      Analyze the following anomaly detected in the SPC (Statistical Process Control) system.
      
      Anomaly Details:
      ID: ${anomaly.id}
      Sample ID: ${anomaly.sampleId}
      Message: ${anomaly.message}
      Timestamp: ${anomaly.timestamp}
      
      Context Data (Recent Measurements):
      ${JSON.stringify(contextData)}

      Provide a concise 3-sentence root cause analysis and 1 specific actionable recommendation.
      Do not use markdown formatting.
    `;

    const response = await ai.models.generateContent({
      model: 'gemini-3-flash-preview',
      contents: prompt,
    });

    return response.text || "No analysis generated.";
  } catch (error) {
    console.error("Gemini Analysis Failed:", error);
    return "Failed to generate AI analysis. Please try again later.";
  }
};