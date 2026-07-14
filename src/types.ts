import { Timestamp } from 'firebase/firestore';

export type PlanStatus = 'active' | 'completed';
export type AssetStatus = 'pending' | 'checked';
export type CheckResult = 'normal' | 'abnormal';

export interface InventoryPlan {
  id?: string;
  name: string;
  description: string;
  scope: string;
  status: PlanStatus;
  createdAt: Timestamp;
}

export interface Asset {
  id?: string;
  planId: string;
  assetCode: string; // Concatenated AssetCode-SubNumber
  companyCode: string;
  accountName: string;
  categoryName: string;
  assetDescription: string;
  acquisitionDate: string;
  acquisitionCost: number;
  bookValue: number;
  quantity: number;
  unit: string;
  costCenter: string;
  originalCustodian: string;
  originalLocation: string;
  originalOffice: string; // "室"
  status: AssetStatus;
  checkResult?: CheckResult;
  checkRemark?: string;
  checkTime?: Timestamp;
  checkBy?: string;
  updatedCustodian?: string;
  updatedLocation?: string;
  updatedOffice?: string;
}

export interface UserProfile {
  uid: string;
  email: string;
  role: 'admin' | 'pda_user';
  name: string;
}
