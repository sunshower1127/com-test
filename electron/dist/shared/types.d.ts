export interface ExecuteRequest {
    type: 'execute';
    id: string;
    code: string;
}
export interface LaunchAppRequest {
    type: 'launch';
    id: string;
    app: 'excel' | 'hwp';
}
export interface QuitAppRequest {
    type: 'quit';
    id: string;
    app: 'excel' | 'hwp';
}
export interface ListMembersRequest {
    type: 'list-members';
    id: string;
    app: 'excel' | 'hwp';
}
export type WorkerRequest = ExecuteRequest | LaunchAppRequest | QuitAppRequest | ListMembersRequest;
export interface ExecuteSuccess {
    type: 'result';
    id: string;
    success: true;
    result: unknown;
    logs: string[];
}
export interface ExecuteError {
    type: 'error';
    id: string;
    success: false;
    error: string;
    stack?: string;
    line?: number;
    code: string;
    logs: string[];
    restored: boolean;
}
export interface AppStatus {
    type: 'status';
    excel: boolean;
    hwp: boolean;
}
export interface MemberList {
    type: 'members';
    id: string;
    members: Array<{
        name: string;
        kind: string;
        params: Array<{
            name: string;
            vt: string;
        }>;
        returnType: string;
    }>;
}
export type WorkerResponse = ExecuteSuccess | ExecuteError | AppStatus | MemberList;
